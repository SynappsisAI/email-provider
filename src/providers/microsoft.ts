import { ConfidentialClientApplication } from "@azure/msal-node";
import type {
  Attachment,
  EmailProvider,
  EmailAddress,
  MessageSummary,
  MessageFull,
  Folder,
  ListResult,
  RetrievedAttachment,
  SendParams,
  ListParams,
  SearchParams,
  ReplyParams,
  ForwardParams,
  MicrosoftProviderConfig,
} from "../types.js";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

export class MicrosoftEmailProvider implements EmailProvider {
  readonly name = "microsoft";
  private msalClient: ConfidentialClientApplication;

  constructor(config: MicrosoftProviderConfig) {
    this.msalClient = new ConfidentialClientApplication({
      auth: {
        clientId: config.clientId,
        clientSecret: config.clientSecret,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
      },
    });
  }

  // ── Graph helpers ──

  private async getToken(): Promise<string> {
    const result = await this.msalClient.acquireTokenByClientCredential({
      scopes: ["https://graph.microsoft.com/.default"],
    });
    if (!result?.accessToken) throw new Error("Failed to acquire Microsoft access token");
    return result.accessToken;
  }

  private async graph(
    path: string,
    opts: { method?: string; body?: unknown; params?: Record<string, string> } = {}
  ) {
    const token = await this.getToken();
    const url = new URL(`${GRAPH_BASE}${path}`);
    if (opts.params) {
      for (const [k, v] of Object.entries(opts.params)) url.searchParams.set(k, v);
    }

    const res = await fetch(url.toString(), {
      method: opts.method ?? "GET",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: opts.body ? JSON.stringify(opts.body) : undefined,
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({ error: { message: res.statusText } }));
      throw new Error(`Graph API ${res.status}: ${(err as any)?.error?.message ?? res.statusText}`);
    }

    const text = await res.text();
    return text ? JSON.parse(text) : null;
  }

  // ── Normalization ──

  private static addr(o: any): EmailAddress {
    return { address: o?.emailAddress?.address ?? "", name: o?.emailAddress?.name };
  }

  private static addrs(arr: any[]): EmailAddress[] {
    return (arr ?? []).map(MicrosoftEmailProvider.addr);
  }

  private static toSummary(m: any): MessageSummary {
    return {
      id: m.id,
      subject: m.subject ?? "",
      from: MicrosoftEmailProvider.addr(m.from),
      to: MicrosoftEmailProvider.addrs(m.toRecipients),
      receivedAt: m.receivedDateTime,
      isRead: m.isRead,
      preview: m.bodyPreview ?? "",
      hasAttachments: m.hasAttachments,
    };
  }

  private static toFull(m: any): MessageFull {
    return {
      id: m.id,
      subject: m.subject ?? "",
      from: MicrosoftEmailProvider.addr(m.from),
      to: MicrosoftEmailProvider.addrs(m.toRecipients),
      cc: MicrosoftEmailProvider.addrs(m.ccRecipients),
      bcc: MicrosoftEmailProvider.addrs(m.bccRecipients),
      replyTo: MicrosoftEmailProvider.addrs(m.replyTo),
      receivedAt: m.receivedDateTime,
      isRead: m.isRead,
      bodyHtml: m.body?.contentType === "html" ? m.body.content : "",
      bodyText: m.body?.contentType === "text" ? m.body.content : "",
      hasAttachments: m.hasAttachments,
      attachments: [],
    };
  }

  // ── EmailProvider implementation ──

  async sendEmail({ mailbox, to, subject, body, cc, bcc, saveToSent, attachments }: SendParams) {
    const message: Record<string, unknown> = {
      subject,
      body: { contentType: "HTML", content: body },
      toRecipients: to.map((a) => ({ emailAddress: { address: a } })),
    };
    if (cc?.length) message.ccRecipients = cc.map((a) => ({ emailAddress: { address: a } }));
    if (bcc?.length) message.bccRecipients = bcc.map((a) => ({ emailAddress: { address: a } }));
    if (attachments?.length) {
      // Graph API requires `contentBytes` as a base64 string. Encode the Buffer
      // here so callers never see base64.
      message.attachments = attachments.map((att) => ({
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: att.filename,
        contentType: att.contentType,
        contentBytes: att.content.toString("base64"),
        ...(att.contentId ? { isInline: true, contentId: att.contentId } : {}),
      }));
    }

    await this.graph(`/users/${mailbox}/sendMail`, {
      method: "POST",
      body: { message, saveToSentItems: saveToSent ?? true },
    });

    return { from: mailbox };
  }

  async listMessages({ mailbox, folder, limit, offset, unreadOnly }: ListParams): Promise<ListResult> {
    const top = Math.min(limit ?? 25, 100);
    const skip = offset ?? 0;
    const folderName = folder ?? "inbox";

    const params: Record<string, string> = {
      $top: String(top),
      $skip: String(skip),
      $orderby: "receivedDateTime desc",
      $select: "id,subject,from,toRecipients,receivedDateTime,isRead,bodyPreview,hasAttachments",
      $count: "true",
    };
    if (unreadOnly) params.$filter = "isRead eq false";

    const r = await this.graph(`/users/${mailbox}/mailFolders/${folderName}/messages`, { params });
    return {
      messages: (r.value as any[]).map(MicrosoftEmailProvider.toSummary),
      totalCount: r["@odata.count"] ?? r.value.length,
    };
  }

  async readMessage(mailbox: string, messageId: string): Promise<MessageFull> {
    const [r, attRes] = await Promise.all([
      this.graph(`/users/${mailbox}/messages/${messageId}`, {
        params: {
          $select:
            "id,subject,from,toRecipients,ccRecipients,bccRecipients,replyTo,receivedDateTime,isRead,body,hasAttachments",
        },
      }),
      this.graph(`/users/${mailbox}/messages/${messageId}/attachments`, {
        params: { $select: "id,name,contentType,size" },
      }).catch(() => ({ value: [] })),
    ]);
    const full = MicrosoftEmailProvider.toFull(r);
    full.attachments = (attRes.value as any[]).map((a: any) => ({
      id: a.id,
      filename: a.name,
      contentType: a.contentType,
      size: a.size,
    }));
    full.hasAttachments = full.attachments.length > 0;
    return full;
  }

  async getAttachment(
    mailbox: string,
    messageId: string,
    attachmentId: string,
  ): Promise<RetrievedAttachment> {
    const r = await this.graph(
      `/users/${mailbox}/messages/${messageId}/attachments/${attachmentId}`
    );
    // Graph returns the file content as base64 in `contentBytes`. Decode to Buffer.
    const content = Buffer.from(r.contentBytes ?? "", "base64");
    return {
      filename: r.name ?? "",
      contentType: r.contentType ?? "",
      size: r.size ?? content.length,
      content,
    };
  }

  async searchMessages({ mailbox, query, folder, limit, offset }: SearchParams): Promise<ListResult> {
    const top = Math.min(limit ?? 25, 100);

    const params: Record<string, string> = {
      $search: `"${query}"`,
      $top: String(top),
      $select: "id,subject,from,toRecipients,receivedDateTime,isRead,bodyPreview,hasAttachments",
    };
    if (offset && offset > 0) params.$skip = String(offset);

    const basePath = folder
      ? `/users/${mailbox}/mailFolders/${folder}/messages`
      : `/users/${mailbox}/messages`;

    const r = await this.graph(basePath, { params });
    return {
      messages: (r.value as any[]).map(MicrosoftEmailProvider.toSummary),
      totalCount: r["@odata.count"] ?? r.value.length,
    };
  }

  async replyToMessage({ mailbox, messageId, body, replyAll, attachments }: ReplyParams) {
    if (attachments?.length) {
      await this.sendViaDraft(mailbox, messageId, replyAll ? "createReplyAll" : "createReply", body, attachments);
      return;
    }
    const action = replyAll ? "replyAll" : "reply";
    await this.graph(`/users/${mailbox}/messages/${messageId}/${action}`, {
      method: "POST",
      body: { comment: body },
    });
  }

  async forwardMessage({ mailbox, messageId, to, comment, attachments }: ForwardParams) {
    if (attachments?.length) {
      await this.sendViaDraft(mailbox, messageId, "createForward", comment ?? "", attachments, to);
      return;
    }
    await this.graph(`/users/${mailbox}/messages/${messageId}/forward`, {
      method: "POST",
      body: {
        comment: comment ?? "",
        toRecipients: to.map((a) => ({ emailAddress: { address: a } })),
      },
    });
  }

  /** Single-request attachment upload limit on the draft path (bigger needs an upload session). */
  private static readonly SIMPLE_ATTACHMENT_MAX_BYTES = 3 * 1024 * 1024;

  /**
   * The original's inline images that a reply draft's quoted body references but that the
   * draft doesn't already carry. BEST-EFFORT: a quoted image that can't be copied (too big
   * for a simple upload, a transient error) must never block the reply itself.
   */
  private async originalInlineAttachments(mailbox: string, messageId: string, draftPath: string): Promise<Attachment[]> {
    try {
      // Metadata only (a full list would pull every attachment's bytes, PDFs included).
      const base = `/users/${mailbox}/messages/${messageId}/attachments`;
      const [list, onDraft] = await Promise.all([
        this.graph(base, { params: { $select: "id,isInline,size" } }),
        this.graph(`${draftPath}/attachments`, { params: { $select: "id,isInline" } }),
      ]);
      const inlineIds = ((list?.value ?? []) as any[])
        .filter((a) => a.isInline && (a.size ?? 0) < MicrosoftEmailProvider.SIMPLE_ATTACHMENT_MAX_BYTES)
        .map((a) => a.id as string);
      if (!inlineIds.length) return [];
      // If Graph already copied the inline images into the draft, don't add them twice.
      const draftInlineIds = ((onDraft?.value ?? []) as any[]).filter((a) => a.isInline).map((a) => a.id as string);
      const draftCids = new Set(
        (await Promise.all(draftInlineIds.map((id) => this.graph(`${draftPath}/attachments/${id}`))))
          .map((a: any) => String(a?.contentId ?? "").replace(/^<|>$/g, "")),
      );
      const full = await Promise.all(inlineIds.map((id) => this.graph(`${base}/${id}`)));
      return full
        .filter((a: any) => a?.contentId && a.contentBytes && a["@odata.type"] === "#microsoft.graph.fileAttachment")
        .map((a: any) => ({
          filename: a.name,
          contentType: a.contentType,
          content: Buffer.from(a.contentBytes, "base64"),
          contentId: String(a.contentId).replace(/^<|>$/g, ""),
        }))
        .filter((a) => !draftCids.has(a.contentId));
    } catch {
      return [];
    }
  }

  /**
   * Reply/forward carrying attachments. The one-shot `/reply` and `/forward` actions only
   * take a `comment` string, so inline images can't ride on them. Instead: create the
   * draft (Graph fills threading, quoted original and — on forward — the original's
   * attachments), prepend our HTML to its body, add the attachments, send. A failure
   * after the draft exists deletes it so no orphan draft is left behind.
   */
  private async sendViaDraft(
    mailbox: string,
    messageId: string,
    action: "createReply" | "createReplyAll" | "createForward",
    html: string,
    attachments: Attachment[],
    to?: string[],
  ) {
    const draft = await this.graph(`/users/${mailbox}/messages/${messageId}/${action}`, { method: "POST", body: {} });
    const draftPath = `/users/${mailbox}/messages/${draft.id}`;
    try {
      // The draft's quoted body follows the mailbox's compose format — normalize a
      // plain-text one to HTML so its line breaks survive under contentType HTML.
      const raw: string = draft.body?.content ?? "";
      const quoted = draft.body?.contentType?.toLowerCase() === "html"
        ? raw
        : raw.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\r?\n/g, "<br>");
      const bodyTag = quoted.match(/<body[^>]*>/i);
      // Slice, not String.replace: a replacement string would expand `$&`/`$'`/`$$` in the HTML.
      const content = bodyTag?.index !== undefined
        ? quoted.slice(0, bodyTag.index + bodyTag[0].length) + html + quoted.slice(bodyTag.index + bodyTag[0].length)
        : `${html}${quoted}`;
      const patch: Record<string, unknown> = { body: { contentType: "HTML", content } };
      if (to) patch.toRecipients = to.map((a) => ({ emailAddress: { address: a } }));
      await this.graph(draftPath, { method: "PATCH", body: patch });
      // createReply(All) drafts quote the original's HTML (incl. its cid: images) but don't
      // carry its inline attachments — copy them so the quote doesn't render broken images.
      // (createForward already carries every original attachment.)
      const carried = action === "createForward" ? [] : await this.originalInlineAttachments(mailbox, messageId, draftPath);
      for (const att of [...carried, ...attachments]) {
        await this.graph(`${draftPath}/attachments`, {
          method: "POST",
          body: {
            "@odata.type": "#microsoft.graph.fileAttachment",
            name: att.filename,
            contentType: att.contentType,
            contentBytes: att.content.toString("base64"),
            ...(att.contentId ? { isInline: true, contentId: att.contentId } : {}),
          },
        });
      }
      await this.graph(`${draftPath}/send`, { method: "POST" });
    } catch (err) {
      await this.graph(draftPath, { method: "DELETE" }).catch(() => {});
      throw err;
    }
  }

  async markAsRead(mailbox: string, messageId: string, isRead: boolean) {
    await this.graph(`/users/${mailbox}/messages/${messageId}`, {
      method: "PATCH",
      body: { isRead },
    });
  }

  async listFolders(mailbox: string, parentId?: string): Promise<Folder[]> {
    const path = parentId
      ? `/users/${mailbox}/mailFolders/${parentId}/childFolders`
      : `/users/${mailbox}/mailFolders`;

    const r = await this.graph(path, {
      params: { $select: "id,displayName,totalItemCount,unreadItemCount,childFolderCount" },
    });

    return (r.value as any[]).map((f) => ({
      id: f.id,
      name: f.displayName,
      totalCount: f.totalItemCount,
      unreadCount: f.unreadItemCount,
      childCount: f.childFolderCount,
    }));
  }
}

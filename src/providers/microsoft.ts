import { randomUUID } from "node:crypto";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { buildMimeMessage } from "../mime.js";
import { assertOriginHeaderName, classifyKind, originHeaderPair, readOrigin } from "../headers.js";
import type {
  Attachment,
  ConversationMessageParams,
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

// Fields every MessageSummary needs.
const SUMMARY_SELECT =
  "id,subject,from,toRecipients,ccRecipients,conversationId,internetMessageId,parentFolderId,receivedDateTime,isRead,bodyPreview,hasAttachments";

// Exchange's message class (PR_MESSAGE_CLASS) — tells OOF, read receipts, NDRs and
// meeting items apart where headers alone are unreliable.
const MESSAGE_CLASS_PROP = "String 0x001A";

/** OData string literal: single quotes doubled. */
const odataString = (v: string) => `'${v.replace(/'/g, "''")}'`;

class GraphError extends Error {
  constructor(readonly status: number, readonly code: string | undefined, message: string) {
    super(message);
  }
}

// Stamps Exchange Online puts on mail submitted by an authenticated sender of the
// same organization. Exchange strips or overwrites inbound X-MS-Exchange-* headers,
// so they can't be forged from outside — unlike From, DKIM or DMARC (internal mail
// is unsigned: dkim=none, dmarc=none).
const INTERNAL_SENDER_STAMPS = new Map([
  ["x-ms-exchange-organization-authas", "internal"],
  ["x-ms-exchange-organization-messagedirectionality", "originating"],
  ["x-ms-exchange-crosstenant-fromentityheader", "hosted"],
]);

// The stamps describe whoever SUBMITTED the message, not the owner of From. A rule
// redirect / auto-forward resubmits an outside message (spoofed From kept) as
// internal, so any sign of resubmission or automatic generation disqualifies it.
const RESUBMISSION_MARKERS = [
  /^x-ms-exchange-organization-autoforwarded$/,
  /^x-ms-exchange-inbox-rules-loop$/,
  /^x-ms-exchange-generated-message-source$/,
  /^resent-/,
];

type GraphRecipient = { emailAddress?: { address?: string } } | null | undefined;

/** Microsoft side of `MessageFull.internalSender`, from a Graph message with `from`,
 *  `sender` and `internetMessageHeaders`. True iff:
 *   - every internal-sender stamp is present exactly once with its expected value;
 *   - `sender` and `from` are the same address (rules out send on behalf);
 *   - no resubmission marker (rule redirect / auto-forward / Resent-*) is present and
 *     `Auto-Submitted` is absent or `no`.
 *  Names and values compare case-insensitively, values trimmed. Anything ambiguous →
 *  false. Exported for tests; not part of the public API. */
export function isExchangeInternalSender(m: {
  from?: GraphRecipient;
  sender?: GraphRecipient;
  internetMessageHeaders?: { name?: string; value?: string }[] | null;
}): boolean {
  const from = (m.from?.emailAddress?.address ?? "").trim().toLowerCase();
  const sender = (m.sender?.emailAddress?.address ?? "").trim().toLowerCase();
  if (!from || from !== sender) return false;

  const seen = new Map<string, string[]>();
  for (const h of m.internetMessageHeaders ?? []) {
    const name = (h?.name ?? "").trim().toLowerCase();
    const value = (h?.value ?? "").trim().toLowerCase();
    if (RESUBMISSION_MARKERS.some((re) => re.test(name))) return false;
    if (name === "auto-submitted" && value !== "no") return false;
    if (INTERNAL_SENDER_STAMPS.has(name)) seen.set(name, [...(seen.get(name) ?? []), value]);
  }
  return [...INTERNAL_SENDER_STAMPS].every(([name, expected]) => {
    const values = seen.get(name) ?? [];
    return values.length === 1 && values[0] === expected;
  });
}

export class MicrosoftEmailProvider implements EmailProvider {
  readonly name = "microsoft";
  private msalClient: ConfidentialClientApplication;
  private immutableIds: boolean;
  private originHeader?: string;

  constructor(config: MicrosoftProviderConfig) {
    this.immutableIds = config.immutableIds ?? false;
    this.originHeader = config.originHeader === undefined ? undefined : assertOriginHeaderName(config.originHeader);
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
    opts: {
      method?: string;
      body?: unknown;
      /** Sent as-is instead of JSON (MIME sendMail). */
      rawBody?: string;
      contentType?: string;
      params?: Record<string, string>;
    } = {}
  ) {
    const token = await this.getToken();
    const url = new URL(`${GRAPH_BASE}${path}`);
    if (opts.params) {
      for (const [k, v] of Object.entries(opts.params)) url.searchParams.set(k, v);
    }

    const headers: Record<string, string> = {
      Authorization: `Bearer ${token}`,
      "Content-Type": opts.contentType ?? "application/json",
    };
    if (this.immutableIds) headers.Prefer = 'IdType="ImmutableId"';

    const res = await fetch(url.toString(), {
      method: opts.method ?? "GET",
      headers,
      body: opts.rawBody ?? (opts.body ? JSON.stringify(opts.body) : undefined),
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({ error: { message: res.statusText } }));
      const e = (err as any)?.error;
      throw new GraphError(res.status, e?.code, `Graph API ${res.status}: ${e?.message ?? res.statusText}`);
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
      cc: MicrosoftEmailProvider.addrs(m.ccRecipients),
      conversationId: m.conversationId ?? undefined,
      internetMessageId: m.internetMessageId ?? undefined,
      parentFolderId: m.parentFolderId ?? undefined,
      receivedAt: m.receivedDateTime,
      isRead: m.isRead,
      preview: m.bodyPreview ?? "",
      hasAttachments: m.hasAttachments,
    };
  }

  private static toFull(m: any, originHeader?: string): MessageFull {
    const headers = m.internetMessageHeaders;
    // Graph echoes the property id normalized ("String 0x1a"), not as requested.
    const messageClass = ((m.singleValueExtendedProperties ?? []) as any[]).find((p) =>
      /^string\s+0x0*1a$/i.test(String(p?.id ?? "").trim()),
    )?.value;
    return {
      id: m.id,
      subject: m.subject ?? "",
      from: MicrosoftEmailProvider.addr(m.from),
      to: MicrosoftEmailProvider.addrs(m.toRecipients),
      cc: MicrosoftEmailProvider.addrs(m.ccRecipients),
      conversationId: m.conversationId ?? undefined,
      internetMessageId: m.internetMessageId ?? undefined,
      parentFolderId: m.parentFolderId ?? undefined,
      uniqueBodyHtml: m.uniqueBody ? (m.uniqueBody.contentType === "html" ? m.uniqueBody.content : "") : undefined,
      uniqueBodyText: m.uniqueBody ? (m.uniqueBody.contentType === "text" ? m.uniqueBody.content : "") : undefined,
      kind: classifyKind({
        headers,
        fromAddress: m.from?.emailAddress?.address,
        messageClass,
        isEventMessage: /eventmessage/i.test(m["@odata.type"] ?? ""),
      }),
      origin: readOrigin(originHeader, headers),
      bcc: MicrosoftEmailProvider.addrs(m.bccRecipients),
      replyTo: MicrosoftEmailProvider.addrs(m.replyTo),
      isMailingList: ((m.internetMessageHeaders ?? []) as any[]).some((h) =>
        /^list-(id|post)$/i.test(h?.name ?? ""),
      ),
      internalSender: isExchangeInternalSender(m),
      receivedAt: m.receivedDateTime,
      isRead: m.isRead,
      bodyHtml: m.body?.contentType === "html" ? m.body.content : "",
      bodyText: m.body?.contentType === "text" ? m.body.content : "",
      hasAttachments: m.hasAttachments,
      attachments: [],
    };
  }

  // ── Neutral folder name → Graph well-known folder name ──

  // Shares the neutral names of GoogleEmailProvider.LABEL_TO_FOLDER, except the
  // Gmail-only "starred"/"important". Anything not listed (real folder ids, Graph
  // well-known names like "archive") passes through untouched.
  private static readonly NEUTRAL_TO_FOLDER: Record<string, string> = {
    inbox: "inbox",
    sent: "sentitems",
    drafts: "drafts",
    draft: "drafts",
    trash: "deleteditems",
    spam: "junkemail",
    junk: "junkemail",
  };

  private static toFolderId(folder: string): string {
    return MicrosoftEmailProvider.NEUTRAL_TO_FOLDER[folder.trim().toLowerCase()] ?? folder;
  }

  // ── EmailProvider implementation ──

  /** Origin as Graph `internetMessageHeaders` — only settable when the message is created. */
  private originHeaders(origin: string | undefined) {
    const pairs = originHeaderPair(this.originHeader, origin);
    return pairs.length ? { internetMessageHeaders: pairs.map(([name, value]) => ({ name, value })) } : {};
  }

  async sendEmail({ mailbox, to, subject, body, cc, bcc, saveToSent, attachments, origin }: SendParams) {
    const message: Record<string, unknown> = {
      subject,
      body: { contentType: "HTML", content: body },
      toRecipients: to.map((a) => ({ emailAddress: { address: a } })),
      ...this.originHeaders(origin),
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
    const folderName = MicrosoftEmailProvider.toFolderId(folder ?? "inbox");

    const params: Record<string, string> = {
      $top: String(top),
      $skip: String(skip),
      $orderby: "receivedDateTime desc",
      $select: SUMMARY_SELECT,
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
            "id,subject,from,sender,toRecipients,ccRecipients,bccRecipients,replyTo,internetMessageHeaders,conversationId,internetMessageId,parentFolderId,uniqueBody,receivedDateTime,isRead,body,hasAttachments",
          $expand: `singleValueExtendedProperties($filter=id eq ${odataString(MESSAGE_CLASS_PROP)})`,
        },
      }),
      this.graph(`/users/${mailbox}/messages/${messageId}/attachments`, {
        params: { $select: "id,name,contentType,size" },
      }).catch(() => ({ value: [] })),
    ]);
    const full = MicrosoftEmailProvider.toFull(r, this.originHeader);
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
      $select: SUMMARY_SELECT,
    };
    if (offset && offset > 0) params.$skip = String(offset);

    const basePath = folder
      ? `/users/${mailbox}/mailFolders/${MicrosoftEmailProvider.toFolderId(folder)}/messages`
      : `/users/${mailbox}/messages`;

    const r = await this.graph(basePath, { params });
    return {
      messages: (r.value as any[]).map(MicrosoftEmailProvider.toSummary),
      totalCount: r["@odata.count"] ?? r.value.length,
    };
  }

  // An origin header can only be set when the message is created, and only the
  // create* (draft) actions were verified to take it — so origin routes via the draft.
  async replyToMessage({ mailbox, messageId, body, replyAll, attachments, origin }: ReplyParams) {
    if (attachments?.length || origin !== undefined) {
      await this.sendViaDraft(mailbox, messageId, replyAll ? "createReplyAll" : "createReply", body, attachments ?? [], undefined, origin);
      return;
    }
    const action = replyAll ? "replyAll" : "reply";
    await this.graph(`/users/${mailbox}/messages/${messageId}/${action}`, {
      method: "POST",
      body: { comment: body },
    });
  }

  async forwardMessage({ mailbox, messageId, to, comment, attachments, origin }: ForwardParams) {
    if (attachments?.length || origin !== undefined) {
      await this.sendViaDraft(mailbox, messageId, "createForward", comment ?? "", attachments ?? [], to, origin);
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

  private async addDraftAttachment(draftPath: string, att: Attachment) {
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
    origin?: string,
  ) {
    const headers = this.originHeaders(origin);
    const draft = await this.graph(`/users/${mailbox}/messages/${messageId}/${action}`, {
      method: "POST",
      body: "internetMessageHeaders" in headers ? { message: headers } : {},
    });
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
      // The original's quoted images are best-effort: a rejected upload just leaves that
      // quoted image broken. The caller's own attachments stay fail-hard.
      for (const att of carried) await this.addDraftAttachment(draftPath, att).catch(() => {});
      for (const att of attachments) await this.addDraftAttachment(draftPath, att);
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

    return (r.value as any[]).map(MicrosoftEmailProvider.toFolder);
  }

  private static toFolder(f: any): Folder {
    return {
      id: f.id,
      name: f.displayName,
      totalCount: f.totalItemCount ?? 0,
      unreadCount: f.unreadItemCount ?? 0,
      childCount: f.childFolderCount ?? 0,
    };
  }

  async sendConversationMessage(p: ConversationMessageParams): Promise<{ messageId: string }> {
    const messageId = `<${randomUUID()}@${p.mailbox.split("@")[1] ?? "localhost"}>`;
    const mime = buildMimeMessage({
      from: p.mailbox,
      to: p.to,
      cc: p.cc,
      subject: p.subject,
      html: p.html,
      text: p.text,
      messageId,
      inReplyTo: p.inReplyTo,
      references: p.references,
      headers: [
        ...originHeaderPair(this.originHeader, p.origin),
        ...(p.autoSubmitted
          ? ([["Auto-Submitted", p.autoSubmitted], ["X-Auto-Response-Suppress", "All"]] as [string, string][])
          : []),
      ],
      attachments: p.attachments,
    });
    // Graph's MIME send: the whole message, standard base64, as text/plain. Graph keeps
    // our Message-ID and threading headers, and saves the copy to Sent Items.
    await this.graph(`/users/${p.mailbox}/sendMail`, {
      method: "POST",
      rawBody: Buffer.from(mime).toString("base64"),
      contentType: "text/plain",
    });
    return { messageId };
  }

  async findByInternetMessageId(mailbox: string, internetMessageId: string, folder?: string): Promise<MessageSummary[]> {
    const path = folder
      ? `/users/${mailbox}/mailFolders/${MicrosoftEmailProvider.toFolderId(folder)}/messages`
      : `/users/${mailbox}/messages`;
    const r = await this.graph(path, {
      params: { $filter: `internetMessageId eq ${odataString(internetMessageId)}`, $select: SUMMARY_SELECT, $top: "10" },
    });
    return (r.value as any[]).map(MicrosoftEmailProvider.toSummary);
  }

  async moveMessage(mailbox: string, messageId: string, destinationFolder: string): Promise<{ id: string }> {
    const r = await this.graph(`/users/${mailbox}/messages/${messageId}/move`, {
      method: "POST",
      body: { destinationId: MicrosoftEmailProvider.toFolderId(destinationFolder) },
    });
    return { id: r.id };
  }

  async getOrCreateFolder(mailbox: string, name: string, opts: { parentId?: string; hidden?: boolean } = {}): Promise<Folder> {
    const base = opts.parentId
      ? `/users/${mailbox}/mailFolders/${MicrosoftEmailProvider.toFolderId(opts.parentId)}/childFolders`
      : `/users/${mailbox}/mailFolders`;
    const find = async () => {
      const r = await this.graph(base, {
        params: {
          $filter: `displayName eq ${odataString(name)}`,
          $select: "id,displayName,totalItemCount,unreadItemCount,childFolderCount",
          // Hidden folders are left out of listings unless asked for.
          includeHiddenFolders: "true",
        },
      });
      const f = (r.value as any[])[0];
      return f ? MicrosoftEmailProvider.toFolder(f) : undefined;
    };
    const existing = await find();
    if (existing) return existing;
    try {
      // isHidden can only be set at creation.
      const f = await this.graph(base, { method: "POST", body: { displayName: name, isHidden: opts.hidden ?? false } });
      return MicrosoftEmailProvider.toFolder(f);
    } catch (err) {
      // Lost a create race: someone made it between our lookup and the POST.
      if (err instanceof GraphError && err.status === 409) {
        const raced = await find();
        if (raced) return raced;
      }
      throw err;
    }
  }
}

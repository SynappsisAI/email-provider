import { gmail as gmailApi, auth, type gmail_v1 } from "@googleapis/gmail";
import { buildMime } from "../mime.js";
import type {
  Attachment,
  EmailProvider,
  EmailAddress,
  MessageSummary,
  MessageFull,
  AttachmentInfo,
  Folder,
  ListResult,
  RetrievedAttachment,
  SendParams,
  ListParams,
  SearchParams,
  ReplyParams,
  ForwardParams,
  GoogleProviderConfig,
} from "../types.js";

export class GoogleEmailProvider implements EmailProvider {
  readonly name = "google";
  private credentials: Record<string, unknown>;
  private gmailClients = new Map<string, gmail_v1.Gmail>();

  constructor(config: GoogleProviderConfig) {
    this.credentials = config.serviceAccountKey;
  }

  /** Get or create a Gmail client for the given delegated user (cached). */
  private getGmail(mailbox: string): gmail_v1.Gmail {
    let client = this.gmailClients.get(mailbox);
    if (client) return client;

    const googleAuth = new auth.GoogleAuth({
      credentials: this.credentials,
      clientOptions: { subject: mailbox },
      scopes: [
        "https://www.googleapis.com/auth/gmail.modify",
        "https://www.googleapis.com/auth/gmail.send",
      ],
    });

    client = gmailApi({ version: "v1", auth: googleAuth });
    this.gmailClients.set(mailbox, client);
    return client;
  }

  // ── Parsing helpers ──

  private static parseAddress(raw: string): EmailAddress {
    const match = raw.trim().match(/^(.*?)\s*<([^<>]+)>$/);
    if (match) {
      const name = match[1].trim().replace(/^"(.*)"$/, "$1");
      return name ? { name, address: match[2].trim() } : { address: match[2].trim() };
    }
    return { address: raw.trim() };
  }

  /** Split on commas OUTSIDE quoted display names and <…> — `"Soporte, Acme" <t@acme.com>`
   *  is one address, not two. Empty entries are dropped. */
  private static parseAddressList(header: string | undefined): EmailAddress[] {
    if (!header) return [];
    const parts: string[] = [];
    let cur = "", inQuotes = false, inAngle = false;
    for (const ch of header) {
      if (ch === '"' && !inAngle) inQuotes = !inQuotes;
      else if (ch === "<" && !inQuotes) inAngle = true;
      else if (ch === ">" && !inQuotes) inAngle = false;
      if (ch === "," && !inQuotes && !inAngle) { parts.push(cur); cur = ""; } else cur += ch;
    }
    parts.push(cur);
    return parts.map(GoogleEmailProvider.parseAddress).filter((a) => a.address);
  }

  private static getHeader(headers: gmail_v1.Schema$MessagePartHeader[], name: string): string {
    return headers?.find((h) => h.name?.toLowerCase() === name.toLowerCase())?.value ?? "";
  }

  private static decodeBody(part: gmail_v1.Schema$MessagePart | undefined): { html: string; text: string } {
    if (!part) return { html: "", text: "" };

    if (part.mimeType === "text/html" && part.body?.data) {
      return { html: Buffer.from(part.body.data, "base64url").toString("utf-8"), text: "" };
    }
    if (part.mimeType === "text/plain" && part.body?.data) {
      return { html: "", text: Buffer.from(part.body.data, "base64url").toString("utf-8") };
    }

    let html = "";
    let text = "";
    for (const sub of part.parts ?? []) {
      const decoded = GoogleEmailProvider.decodeBody(sub);
      if (decoded.html) html = decoded.html;
      if (decoded.text) text = decoded.text;
    }
    return { html, text };
  }

  private static toSummary(m: gmail_v1.Schema$Message): MessageSummary {
    const headers = m.payload?.headers ?? [];
    return {
      id: m.id!,
      subject: GoogleEmailProvider.getHeader(headers, "Subject"),
      from: GoogleEmailProvider.parseAddress(GoogleEmailProvider.getHeader(headers, "From")),
      to: GoogleEmailProvider.parseAddressList(GoogleEmailProvider.getHeader(headers, "To")),
      receivedAt: new Date(Number(m.internalDate)).toISOString(),
      isRead: !m.labelIds?.includes("UNREAD"),
      preview: m.snippet ?? "",
      hasAttachments: m.payload?.parts?.some((p) => p.filename && p.filename.length > 0) ?? false,
    };
  }

  private static extractAttachments(part: gmail_v1.Schema$MessagePart | undefined): AttachmentInfo[] {
    if (!part) return [];
    const results: AttachmentInfo[] = [];
    if (part.filename && part.filename.length > 0 && part.body?.attachmentId) {
      results.push({
        id: part.body.attachmentId,
        filename: part.filename,
        contentType: part.mimeType ?? "application/octet-stream",
        size: part.body.size ?? 0,
      });
    }
    for (const sub of part.parts ?? []) {
      results.push(...GoogleEmailProvider.extractAttachments(sub));
    }
    return results;
  }

  /** Gmail caps a message at 25MB; stay under it with the base64-encoded payload. */
  private static readonly FORWARD_CARRY_MAX_BYTES = 24 * 1024 * 1024;

  /**
   * Attachment parts with what a re-send needs. A part stays INLINE (keeps its Content-ID)
   * only when the html actually references it as `cid:` or it is marked
   * `Content-Disposition: inline` — Outlook stamps a Content-ID on ordinary PDFs too.
   * Inline parts without a filename (common for pasted images) are kept under a fallback
   * name, or the forwarded html's cid: refs would break.
   */
  private static collectParts(
    part: gmail_v1.Schema$MessagePart | undefined,
    html: string,
  ): { attachmentId: string; filename: string; contentType: string; size: number; contentId?: string }[] {
    if (!part) return [];
    const results: { attachmentId: string; filename: string; contentType: string; size: number; contentId?: string }[] = [];
    if (part.body?.attachmentId) {
      const headers = part.headers ?? [];
      const cid = GoogleEmailProvider.getHeader(headers, "Content-ID").replace(/^<|>$/g, "");
      const disposition = GoogleEmailProvider.getHeader(headers, "Content-Disposition").toLowerCase();
      const inline = !!cid && (html.includes(`cid:${cid}`) || disposition.startsWith("inline"));
      if (part.filename || inline) {
        const contentType = part.mimeType ?? "application/octet-stream";
        results.push({
          attachmentId: part.body.attachmentId,
          filename: part.filename || `inline-${part.partId ?? "0"}.${contentType.split("/")[1] ?? "bin"}`,
          contentType,
          size: part.body.size ?? 0,
          ...(inline ? { contentId: cid } : {}),
        });
      }
    }
    for (const sub of part.parts ?? []) results.push(...GoogleEmailProvider.collectParts(sub, html));
    return results;
  }

  private static toFull(m: gmail_v1.Schema$Message): MessageFull {
    const headers = m.payload?.headers ?? [];
    const { html, text } = GoogleEmailProvider.decodeBody(m.payload ?? undefined);
    const attachments = GoogleEmailProvider.extractAttachments(m.payload ?? undefined);
    return {
      id: m.id!,
      subject: GoogleEmailProvider.getHeader(headers, "Subject"),
      from: GoogleEmailProvider.parseAddress(GoogleEmailProvider.getHeader(headers, "From")),
      to: GoogleEmailProvider.parseAddressList(GoogleEmailProvider.getHeader(headers, "To")),
      cc: GoogleEmailProvider.parseAddressList(GoogleEmailProvider.getHeader(headers, "Cc")),
      bcc: GoogleEmailProvider.parseAddressList(GoogleEmailProvider.getHeader(headers, "Bcc")),
      replyTo: GoogleEmailProvider.parseAddressList(GoogleEmailProvider.getHeader(headers, "Reply-To")),
      isMailingList: !!(GoogleEmailProvider.getHeader(headers, "List-Id") || GoogleEmailProvider.getHeader(headers, "List-Post")),
      receivedAt: new Date(Number(m.internalDate)).toISOString(),
      isRead: !m.labelIds?.includes("UNREAD"),
      bodyHtml: html,
      bodyText: text,
      hasAttachments: attachments.length > 0,
      attachments,
    };
  }

  // ── Gmail label ↔ "folder" mapping ──

  private static readonly LABEL_TO_FOLDER: Record<string, string> = {
    inbox: "INBOX",
    sent: "SENT",
    drafts: "DRAFT",
    trash: "TRASH",
    spam: "SPAM",
    starred: "STARRED",
    important: "IMPORTANT",
  };

  private static toLabelId(folder: string): string {
    return GoogleEmailProvider.LABEL_TO_FOLDER[folder.toLowerCase()] ?? folder;
  }

  // ── EmailProvider implementation ──

  async sendEmail({ mailbox, to, subject, body, cc, bcc, attachments }: SendParams) {
    const gmail = this.getGmail(mailbox);
    const raw = buildMime({ from: mailbox, to, cc, bcc, subject, html: body, attachments });

    await gmail.users.messages.send({
      userId: "me",
      requestBody: { raw },
    });

    return { from: mailbox };
  }

  async listMessages({ mailbox, folder, limit, offset, unreadOnly }: ListParams): Promise<ListResult> {
    const gmail = this.getGmail(mailbox);
    const maxResults = Math.min(limit ?? 25, 100);
    const labelIds = [GoogleEmailProvider.toLabelId(folder ?? "inbox")];

    const qParts: string[] = [];
    if (unreadOnly) qParts.push("is:unread");

    const list = await gmail.users.messages.list({
      userId: "me",
      labelIds,
      q: qParts.length ? qParts.join(" ") : undefined,
      maxResults: maxResults + (offset ?? 0),
    });

    const allIds = list.data.messages ?? [];
    const sliced = allIds.slice(offset ?? 0, (offset ?? 0) + maxResults);

    const messages = await Promise.all(
      sliced.map(async ({ id }) => {
        const msg = await gmail.users.messages.get({
          userId: "me",
          id: id!,
          format: "metadata",
          metadataHeaders: ["From", "To", "Subject"],
        });
        return GoogleEmailProvider.toSummary(msg.data);
      })
    );

    return {
      messages,
      totalCount: list.data.resultSizeEstimate ?? messages.length,
    };
  }

  async readMessage(mailbox: string, messageId: string): Promise<MessageFull> {
    const gmail = this.getGmail(mailbox);
    const msg = await gmail.users.messages.get({
      userId: "me",
      id: messageId,
      format: "full",
    });
    return GoogleEmailProvider.toFull(msg.data);
  }

  async searchMessages({ mailbox, query, folder, limit, offset }: SearchParams): Promise<ListResult> {
    const gmail = this.getGmail(mailbox);
    const maxResults = Math.min(limit ?? 25, 100);
    const labelIds = folder ? [GoogleEmailProvider.toLabelId(folder)] : undefined;

    const list = await gmail.users.messages.list({
      userId: "me",
      q: query,
      labelIds,
      maxResults: maxResults + (offset ?? 0),
    });

    const allIds = list.data.messages ?? [];
    const sliced = allIds.slice(offset ?? 0, (offset ?? 0) + maxResults);

    const messages = await Promise.all(
      sliced.map(async ({ id }) => {
        const msg = await gmail.users.messages.get({
          userId: "me",
          id: id!,
          format: "metadata",
          metadataHeaders: ["From", "To", "Subject"],
        });
        return GoogleEmailProvider.toSummary(msg.data);
      })
    );

    return {
      messages,
      totalCount: list.data.resultSizeEstimate ?? messages.length,
    };
  }

  async replyToMessage({ mailbox, messageId, body, replyAll, attachments }: ReplyParams) {
    const gmail = this.getGmail(mailbox);

    const original = await gmail.users.messages.get({
      userId: "me",
      id: messageId,
      format: "metadata",
      metadataHeaders: ["From", "Reply-To", "To", "Cc", "Subject", "Message-ID", "References"],
    });

    const headers = original.data.payload?.headers ?? [];
    const origFrom = GoogleEmailProvider.getHeader(headers, "From");
    // A reply goes to Reply-To when the sender set one (RFC 5322 §3.6.2), else to From —
    // the same rule Graph's native reply and every mail client follow.
    const origReplyTo = GoogleEmailProvider.getHeader(headers, "Reply-To");
    const replyToAddrs = GoogleEmailProvider.parseAddressList(origReplyTo).map((a) => a.address);
    const primary = replyToAddrs.length
      ? replyToAddrs
      : GoogleEmailProvider.parseAddressList(origFrom).map((a) => a.address);
    const origTo = GoogleEmailProvider.getHeader(headers, "To");
    const origCc = GoogleEmailProvider.getHeader(headers, "Cc");
    const origSubject = GoogleEmailProvider.getHeader(headers, "Subject");
    const origMessageId = GoogleEmailProvider.getHeader(headers, "Message-ID");
    const origReferences = GoogleEmailProvider.getHeader(headers, "References");

    const replyTo = replyAll
      ? [...primary, ...GoogleEmailProvider.parseAddressList(origTo).map((a) => a.address)]
          .filter((a) => a.toLowerCase() !== mailbox.toLowerCase())
      : primary;

    const cc = replyAll
      ? GoogleEmailProvider.parseAddressList(origCc)
          .map((a) => a.address)
          .filter((a) => a.toLowerCase() !== mailbox.toLowerCase())
      : undefined;

    const subject = origSubject.startsWith("Re:") ? origSubject : `Re: ${origSubject}`;

    const raw = buildMime({
      from: mailbox,
      to: replyTo,
      cc,
      subject,
      html: body,
      inReplyTo: origMessageId,
      references: origReferences ? `${origReferences} ${origMessageId}` : origMessageId,
      attachments,
    });

    await gmail.users.messages.send({
      userId: "me",
      requestBody: { raw, threadId: original.data.threadId ?? undefined },
    });
  }

  async forwardMessage({ mailbox, messageId, to, comment, attachments }: ForwardParams) {
    const gmail = this.getGmail(mailbox);

    const fullMsg = await gmail.users.messages.get({
      userId: "me",
      id: messageId,
      format: "full",
    });

    const headers = fullMsg.data.payload?.headers ?? [];
    const origSubject = GoogleEmailProvider.getHeader(headers, "Subject");
    const { html, text } = GoogleEmailProvider.decodeBody(fullMsg.data.payload ?? undefined);
    const originalContent = html || `<pre>${text}</pre>`;

    const forwardBody = comment
      ? `${comment}<br><br>---------- Forwarded message ----------<br>${originalContent}`
      : `---------- Forwarded message ----------<br>${originalContent}`;

    const subject = origSubject.startsWith("Fwd:") ? origSubject : `Fwd: ${origSubject}`;

    // Carry the original's attachments (previously dropped), keeping inline images inline
    // so the forwarded body's own cid: references still resolve. Over the cap (Gmail sends
    // ≤25MB of attachments, and the MIME is held in memory ~3x) fall back to the old
    // body-only forward rather than fail.
    const parts = GoogleEmailProvider.collectParts(fullMsg.data.payload ?? undefined, originalContent);
    // Budget the ENCODED size (base64 ≈ 4/3) of the originals plus the caller's own attachments.
    const decoded = parts.reduce((n, p) => n + p.size, 0) + (attachments ?? []).reduce((n, a) => n + a.content.length, 0);
    const carryOriginals = Math.ceil((decoded * 4) / 3) <= GoogleEmailProvider.FORWARD_CARRY_MAX_BYTES;
    const original: Attachment[] = [];
    if (carryOriginals) {
      for (const p of parts) {
        original.push({
          filename: p.filename,
          contentType: p.contentType,
          content: (await this.getAttachment(mailbox, messageId, p.attachmentId)).content,
          ...(p.contentId ? { contentId: p.contentId } : {}),
        });
      }
    }
    const all = [...original, ...(attachments ?? [])];
    const raw = buildMime({ from: mailbox, to, subject, html: forwardBody, attachments: all.length ? all : undefined });

    await gmail.users.messages.send({
      userId: "me",
      requestBody: { raw },
    });
  }

  async markAsRead(mailbox: string, messageId: string, isRead: boolean) {
    const gmail = this.getGmail(mailbox);
    if (isRead) {
      await gmail.users.messages.modify({
        userId: "me",
        id: messageId,
        requestBody: { removeLabelIds: ["UNREAD"] },
      });
    } else {
      await gmail.users.messages.modify({
        userId: "me",
        id: messageId,
        requestBody: { addLabelIds: ["UNREAD"] },
      });
    }
  }

  async getAttachment(
    mailbox: string,
    messageId: string,
    attachmentId: string,
  ): Promise<RetrievedAttachment> {
    const gmail = this.getGmail(mailbox);
    const res = await gmail.users.messages.attachments.get({
      userId: "me",
      messageId,
      id: attachmentId,
    });
    // Gmail returns base64url-encoded data; decode straight to a Buffer.
    const content = Buffer.from(res.data.data ?? "", "base64url");
    return {
      // Gmail's attachment endpoint doesn't return filename or contentType.
      // Caller should fall back to the AttachmentInfo from the parent MessageFull.
      filename: "",
      contentType: "",
      size: res.data.size ?? content.length,
      content,
    };
  }

  async listFolders(mailbox: string): Promise<Folder[]> {
    const gmail = this.getGmail(mailbox);
    const res = await gmail.users.labels.list({ userId: "me" });
    const labels = res.data.labels ?? [];

    const folders = await Promise.all(
      labels.map(async (label) => {
        const detail = await gmail.users.labels.get({
          userId: "me",
          id: label.id!,
        });
        return {
          id: detail.data.id!,
          name: detail.data.name!,
          totalCount: detail.data.messagesTotal ?? 0,
          unreadCount: detail.data.messagesUnread ?? 0,
          childCount: 0,
        };
      })
    );

    return folders;
  }
}

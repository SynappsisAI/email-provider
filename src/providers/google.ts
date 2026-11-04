import { gmail as gmailApi, auth, type gmail_v1 } from "@googleapis/gmail";
import { buildMime, newMessageId } from "../mime.js";
import { decodeCursor, encodeCursor, nextCursor, pageSizeOf, parseSince } from "../cursor.js";
import { assertOriginHeaderName, classifyKind, originHeaderPair, readOrigin } from "../headers.js";
import type {
  Attachment,
  ConversationMessageParams,
  DraftParams,
  DraftRef,
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
  ListSinceParams,
  ListSinceResult,
  SearchParams,
  ReplyParams,
  ForwardParams,
  GoogleProviderConfig,
} from "../types.js";

/** HTTP status of a googleapis/gaxios error. gaxios 7 puts it in `status`; `code` may be a
 *  string (e.g. ERR_BAD_REQUEST) or, in older shapes, the status number. */
function httpStatus(err: any): number | undefined {
  return [err?.status, err?.response?.status, err?.code].find((v) => typeof v === "number");
}

export class GoogleEmailProvider implements EmailProvider {
  readonly name = "google";
  private credentials: Record<string, unknown>;
  private gmailClients = new Map<string, gmail_v1.Gmail>();
  private originHeader?: string;

  constructor(config: GoogleProviderConfig) {
    this.credentials = config.serviceAccountKey;
    this.originHeader = config.originHeader === undefined ? undefined : assertOriginHeaderName(config.originHeader);
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
    // Unbalanced quote/bracket → the "protected" commas are a lie; split on every comma instead
    // (the pre-0.4.1 behavior) so a malformed `<evil@x, ok@y` can't hide an address inside one
    // token. Consumers must still validate each address strictly.
    const list = inQuotes || inAngle ? header.split(",") : parts;
    return list.map(GoogleEmailProvider.parseAddress).filter((a) => a.address);
  }

  /** `list` without duplicates (case-insensitive), entries without an `@`, or anything in
   *  `exclude`, first occurrence kept. */
  private static uniqueAddresses(list: string[], exclude: string[]): string[] {
    const seen = new Set(exclude.map((a) => a.toLowerCase()));
    const out: string[] = [];
    for (const a of list) {
      const key = a.toLowerCase();
      if (a.includes("@") && !seen.has(key)) {
        seen.add(key);
        out.push(a);
      }
    }
    return out;
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

  /** Headers a MessageSummary is built from (metadata-format gets). */
  private static readonly SUMMARY_HEADERS = ["From", "To", "Cc", "Subject", "Message-ID"];

  private static toSummary(m: gmail_v1.Schema$Message): MessageSummary {
    const headers = m.payload?.headers ?? [];
    return {
      id: m.id!,
      subject: GoogleEmailProvider.getHeader(headers, "Subject"),
      from: GoogleEmailProvider.parseAddress(GoogleEmailProvider.getHeader(headers, "From")),
      to: GoogleEmailProvider.parseAddressList(GoogleEmailProvider.getHeader(headers, "To")),
      cc: GoogleEmailProvider.parseAddressList(GoogleEmailProvider.getHeader(headers, "Cc")),
      conversationId: m.threadId ?? undefined,
      internetMessageId: GoogleEmailProvider.getHeader(headers, "Message-ID") || undefined,
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

  /** A calendar invite's text/calendar body part — not an attached .ics file. */
  private static hasInvitePart(part: gmail_v1.Schema$MessagePart | undefined): boolean {
    if (!part) return false;
    const disposition = GoogleEmailProvider.getHeader(part.headers ?? [], "Content-Disposition").toLowerCase();
    const isInvite = part.mimeType?.toLowerCase() === "text/calendar" && !part.filename && !disposition.startsWith("attachment");
    return isInvite || (part.parts ?? []).some(GoogleEmailProvider.hasInvitePart);
  }

  private static toFull(m: gmail_v1.Schema$Message, originHeader?: string): MessageFull {
    const headers = m.payload?.headers ?? [];
    const { html, text } = GoogleEmailProvider.decodeBody(m.payload ?? undefined);
    const attachments = GoogleEmailProvider.extractAttachments(m.payload ?? undefined);
    const from = GoogleEmailProvider.parseAddress(GoogleEmailProvider.getHeader(headers, "From"));
    return {
      id: m.id!,
      subject: GoogleEmailProvider.getHeader(headers, "Subject"),
      from,
      to: GoogleEmailProvider.parseAddressList(GoogleEmailProvider.getHeader(headers, "To")),
      cc: GoogleEmailProvider.parseAddressList(GoogleEmailProvider.getHeader(headers, "Cc")),
      conversationId: m.threadId ?? undefined,
      internetMessageId: GoogleEmailProvider.getHeader(headers, "Message-ID") || undefined,
      kind: classifyKind({
        headers,
        fromAddress: from.address,
        hasCalendarPart: GoogleEmailProvider.hasInvitePart(m.payload ?? undefined),
      }),
      origin: readOrigin(originHeader, headers),
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
    draft: "DRAFT",
    trash: "TRASH",
    spam: "SPAM",
    junk: "SPAM",
    starred: "STARRED",
    important: "IMPORTANT",
  };

  private static toLabelId(folder: string): string {
    return GoogleEmailProvider.LABEL_TO_FOLDER[folder.trim().toLowerCase()] ?? folder;
  }

  // ── EmailProvider implementation ──

  /** MIME of a new message — shared by `sendEmail` and `createDraft`. */
  private newMessageMime({ mailbox, to, subject, body, cc, bcc, attachments, origin }: DraftParams, messageId?: string): string {
    const headers = originHeaderPair(this.originHeader, origin);
    return buildMime({ from: mailbox, to, cc, bcc, subject, html: body, attachments, headers, messageId });
  }

  async sendEmail(params: SendParams) {
    const { mailbox } = params;
    const gmail = this.getGmail(mailbox);
    const raw = this.newMessageMime(params);

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
          metadataHeaders: GoogleEmailProvider.SUMMARY_HEADERS,
        });
        return GoogleEmailProvider.toSummary(msg.data);
      })
    );

    return {
      messages,
      totalCount: list.data.resultSizeEstimate ?? messages.length,
    };
  }

  /** Max concurrent `messages.get` calls (5 quota units each) while building a page. */
  private static readonly GET_CONCURRENCY = 10;
  /** Hard stop for one second holding more messages than this (never seen in practice). */
  private static readonly MAX_IDS_PER_SECOND = 5000;

  /** Ids matching `q` with `labelIds`, up to `cap`; `more` when Gmail has further pages. */
  private async listIds(gmail: gmail_v1.Gmail, q: string, labelIds: string[], cap: number, allPages = false) {
    const ids: string[] = [];
    let pageToken: string | undefined;
    do {
      const r = await gmail.users.messages.list({ userId: "me", q, labelIds, maxResults: Math.min(cap, 500), pageToken });
      ids.push(...(r.data.messages ?? []).map((m) => m.id!));
      pageToken = r.data.nextPageToken ?? undefined;
      if (allPages && ids.length > GoogleEmailProvider.MAX_IDS_PER_SECOND) {
        throw new Error("listSince: too many messages in one second to page through");
      }
    } while (allPages && pageToken);
    return { ids, more: !!pageToken };
  }

  /**
   * Gmail lists newest-first without dates and with no guaranteed order, and `after:` /
   * `before:` only have second granularity. So: find a time window [lo, hi) seconds
   * starting at the cursor that holds few enough ids (halving `hi` while Gmail reports
   * more pages), fetch those messages' metadata, keep exactly [cursor, hi) by
   * `internalDate`, sort ascending, and return the oldest page. An empty window moves on
   * to the rest of the range. Keyset cursor (see cursor.ts) — never an offset.
   */
  async listSince({ mailbox, since, folder, unreadOnly, limit, pageToken }: ListSinceParams): Promise<ListSinceResult> {
    const gmail = this.getGmail(mailbox);
    const pageSize = pageSizeOf(limit);
    const prev = pageToken ? decodeCursor(pageToken, "google") : undefined;
    let fromMs = prev ? prev.t : parseSince(since);
    let seen = new Set(prev?.ids ?? []);
    const labelIds = [GoogleEmailProvider.toLabelId(folder ?? "inbox")];
    const cap = Math.min(Math.max(pageSize * 2, 50), 500);
    // Exclusive end of the whole walk, in seconds, past anything already received.
    const end = Math.floor(Date.now() / 1000) + 2;

    let lo = Math.floor(fromMs / 1000);
    let hi = end;
    // Smallest end seen to overflow from the current lo. An empty [lo, hi) means everything
    // that overflowed lies in [hi, overflowHi), so the search resumes there instead of
    // restarting from `end` — keeps the whole search O(log range) list calls.
    let overflowHi: number | undefined;
    for (let guard = 0; guard < 200; guard++) {
      // ±1 s margins absorb after:/before: inclusiveness; the exact bounds are applied below.
      const q = `after:${lo - 1} before:${hi + 1}${unreadOnly ? " is:unread" : ""}`;
      let { ids, more } = await this.listIds(gmail, q, labelIds, cap);
      if (more && hi - lo > 1) {
        overflowHi = hi;
        hi = lo + Math.floor((hi - lo) / 2);
        continue;
      }
      if (more) ids = (await this.listIds(gmail, q, labelIds, 500, true)).ids;

      const window: MessageSummary[] = [];
      for (let i = 0; i < ids.length; i += GoogleEmailProvider.GET_CONCURRENCY) {
        const got = await Promise.all(
          ids.slice(i, i + GoogleEmailProvider.GET_CONCURRENCY).map((id) =>
            gmail.users.messages
              .get({ userId: "me", id, format: "metadata", metadataHeaders: GoogleEmailProvider.SUMMARY_HEADERS })
              .then((r) => r.data)
              // Deleted between list and get: it's gone, skip it.
              .catch((err: any) => (httpStatus(err) === 404 ? undefined : Promise.reject(err))),
          ),
        );
        for (const m of got) {
          if (!m) continue;
          const t = Number(m.internalDate);
          if (t >= fromMs && t < hi * 1000 && !seen.has(m.id!)) window.push(GoogleEmailProvider.toSummary(m));
        }
      }
      const time = (m: MessageSummary) => Date.parse(m.receivedAt);
      window.sort((a, b) => time(a) - time(b) || (a.id < b.id ? -1 : a.id > b.id ? 1 : 0));

      if (!window.length) {
        if (hi === end) return { messages: [] };
        // Nothing left in this window: continue with the rest of the range.
        fromMs = hi * 1000;
        seen = new Set();
        lo = hi;
        hi = overflowHi ?? end;
        overflowHi = undefined;
        continue;
      }
      const page = window.slice(0, pageSize);
      if (window.length <= pageSize && hi === end) return { messages: page };
      return { messages: page, nextPageToken: encodeCursor(nextCursor("google", page, time, (m) => m.id, prev)) };
    }
    throw new Error("listSince: could not settle on a time window");
  }

  async readMessage(mailbox: string, messageId: string): Promise<MessageFull> {
    const gmail = this.getGmail(mailbox);
    const msg = await gmail.users.messages.get({
      userId: "me",
      id: messageId,
      format: "full",
    });
    return GoogleEmailProvider.toFull(msg.data, this.originHeader);
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
          metadataHeaders: GoogleEmailProvider.SUMMARY_HEADERS,
        });
        return GoogleEmailProvider.toSummary(msg.data);
      })
    );

    return {
      messages,
      totalCount: list.data.resultSizeEstimate ?? messages.length,
    };
  }

  async replyToMessage(params: ReplyParams) {
    const { raw, threadId } = await this.replyMime(params);
    await this.getGmail(params.mailbox).users.messages.send({
      userId: "me",
      requestBody: { raw, threadId },
    });
  }

  /** MIME of a reply / reply-all and the thread to file it in — shared by
   *  `replyToMessage` and `createReplyDraft`. */
  private async replyMime(
    { mailbox, messageId, body, replyAll, attachments, origin }: ReplyParams,
    ownMessageId?: string,
  ): Promise<{ raw: string; threadId?: string }> {
    const gmail = this.getGmail(mailbox);
    const extraHeaders = originHeaderPair(this.originHeader, origin);

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
    // Unusable Reply-To entries (no `@`, e.g. a bare list name) are dropped first, so From
    // is still the fallback when none is left.
    const replyToAddrs = GoogleEmailProvider.parseAddressList(origReplyTo)
      .map((a) => a.address)
      .filter((a) => a.includes("@"));
    // From is ONE mailbox: parse it exactly as readMessage/toFull does (parseAddress, not the
    // comma-splitting list parser) so the address a consumer validated from MessageFull.from is
    // the only one the reply goes to — `From: evil@x, <ok@y>` must not fan out to both.
    const primary = replyToAddrs.length
      ? replyToAddrs
      : [GoogleEmailProvider.parseAddress(origFrom).address].filter(Boolean);
    const origTo = GoogleEmailProvider.getHeader(headers, "To");
    const origCc = GoogleEmailProvider.getHeader(headers, "Cc");
    const origSubject = GoogleEmailProvider.getHeader(headers, "Subject");
    const origMessageId = GoogleEmailProvider.getHeader(headers, "Message-ID");
    const origReferences = GoogleEmailProvider.getHeader(headers, "References");

    // Each address once (case-insensitive); To wins over Cc. Reply-all leaves the mailbox
    // itself out; a plain reply goes to `primary` even when that is the mailbox (a
    // follow-up to its own sent message).
    // Only real addresses: `MAILER-DAEMON`, `undisclosed-recipients:;` and the like are dropped.
    const exclude = replyAll ? [mailbox] : [];
    const replyTo = GoogleEmailProvider.uniqueAddresses(
      replyAll ? [...primary, ...GoogleEmailProvider.parseAddressList(origTo).map((a) => a.address)] : primary,
      exclude,
    );
    let cc = replyAll
      ? GoogleEmailProvider.uniqueAddresses(GoogleEmailProvider.parseAddressList(origCc).map((a) => a.address), [...exclude, ...replyTo])
      : undefined;
    // Only Cc left (e.g. reply-all to the mailbox's note to itself): promote it to To rather
    // than write an empty To header.
    if (!replyTo.length && cc?.length) {
      replyTo.push(...cc);
      cc = undefined;
    }
    // messages.send would reject a reply with no recipient; drafts.create would save it
    // silently, so refuse here for both.
    if (!replyTo.length) throw new Error(`reply to ${messageId}: no recipient (no usable Reply-To/From)`);

    // Case-insensitive: Outlook writes "RE:".
    const subject = /^re:/i.test(origSubject) ? origSubject : `Re: ${origSubject}`;

    const raw = buildMime({
      from: mailbox,
      to: replyTo,
      cc,
      subject,
      html: body,
      messageId: ownMessageId,
      inReplyTo: origMessageId,
      references: origReferences ? `${origReferences} ${origMessageId}` : origMessageId,
      headers: extraHeaders,
      attachments,
    });

    return { raw, threadId: original.data.threadId ?? undefined };
  }

  async forwardMessage(params: ForwardParams) {
    const raw = await this.forwardMime(params);
    await this.getGmail(params.mailbox).users.messages.send({
      userId: "me",
      requestBody: { raw },
    });
  }

  /** MIME of a forward (comment above the original, original's attachments carried) —
   *  shared by `forwardMessage` and `createForwardDraft`. */
  private async forwardMime(
    { mailbox, messageId, to, comment, attachments, origin }: ForwardParams,
    ownMessageId?: string,
  ): Promise<string> {
    const gmail = this.getGmail(mailbox);
    const extraHeaders = originHeaderPair(this.originHeader, origin);

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

    const subject = /^fwd?:/i.test(origSubject) ? origSubject : `Fwd: ${origSubject}`;

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
    return buildMime({
      from: mailbox,
      to,
      subject,
      html: forwardBody,
      messageId: ownMessageId,
      headers: extraHeaders,
      attachments: all.length ? all : undefined,
    });
  }

  // ── Drafts ──
  // Same MIME as the matching send, handed to drafts.create instead of messages.send.
  // Nothing on these paths calls messages.send or drafts.send. Each draft gets our own
  // Message-ID so the caller always has one to find it by.

  async createDraft(params: DraftParams): Promise<DraftRef> {
    const ownId = newMessageId(params.mailbox);
    return this.saveDraft(params.mailbox, this.newMessageMime(params, ownId), ownId);
  }

  async createReplyDraft(params: ReplyParams): Promise<DraftRef> {
    const ownId = newMessageId(params.mailbox);
    const { raw, threadId } = await this.replyMime(params, ownId);
    return this.saveDraft(params.mailbox, raw, ownId, threadId);
  }

  async createForwardDraft(params: ForwardParams): Promise<DraftRef> {
    const ownId = newMessageId(params.mailbox);
    return this.saveDraft(params.mailbox, await this.forwardMime(params, ownId), ownId);
  }

  private async saveDraft(mailbox: string, raw: string, ownId: string, threadId?: string): Promise<DraftRef> {
    const gmail = this.getGmail(mailbox);
    const created = await gmail.users.drafts.create({
      userId: "me",
      requestBody: { message: { raw, threadId } },
    });
    const id = created.data.message?.id;
    if (!created.data.id || !id) {
      // Don't leave a draft behind that the caller can't see or delete.
      if (created.data.id) await gmail.users.drafts.delete({ userId: "me", id: created.data.id }).catch(() => {});
      throw new Error("Gmail drafts.create returned no draft or message id");
    }
    return {
      id,
      draftId: created.data.id,
      // drafts.create keeps the Message-ID we wrote (Gmail may replace it on edit/send).
      internetMessageId: ownId,
      conversationId: created.data.message?.threadId ?? undefined,
    };
  }

  /** drafts.delete only knows draft ids, so a received/sent message's id can't hit it:
   *  Gmail answers 404 (or 400 `invalidArgument` for an id it can't parse) and nothing is
   *  deleted. Any other 400 (e.g. an auth/delegation failure) propagates as is. Gmail
   *  deletes a draft permanently. */
  async deleteDraft(mailbox: string, draftId: string): Promise<void> {
    if (!draftId.trim()) throw new Error("deleteDraft: draftId is empty");
    try {
      await this.getGmail(mailbox).users.drafts.delete({ userId: "me", id: draftId });
    } catch (err: any) {
      const status = httpStatus(err);
      const list = err?.response?.data?.error?.errors ?? err?.errors;
      const reasons: unknown[] = Array.isArray(list) ? list.map((e: any) => e?.reason) : [];
      // invalidArgument is Gmail's generic 400 reason; only its "Invalid id value" means the
      // id isn't a draft id.
      const badId = status === 400 && reasons.includes("invalidArgument") && /invalid\b.*\bid\b/i.test(String(err?.message ?? ""));
      if (status === 404 || badId) {
        // Keep the HTTP status on the error, as Microsoft's GraphError does.
        throw Object.assign(
          new Error(`deleteDraft: no draft ${draftId} in ${mailbox} — already sent or discarded, or not a draft id (pass DraftRef.draftId)`, { cause: err }),
          { status },
        );
      }
      throw err;
    }
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

  /** Message ids matching a Gmail search, as summaries. */
  private async summariesFor(mailbox: string, q: string, labelIds?: string[]): Promise<MessageSummary[]> {
    const gmail = this.getGmail(mailbox);
    const list = await gmail.users.messages.list({ userId: "me", q, labelIds, maxResults: 10 });
    return Promise.all(
      (list.data.messages ?? []).map(async ({ id }) => {
        const msg = await gmail.users.messages.get({
          userId: "me",
          id: id!,
          format: "metadata",
          metadataHeaders: GoogleEmailProvider.SUMMARY_HEADERS,
        });
        return GoogleEmailProvider.toSummary(msg.data);
      }),
    );
  }

  async findByInternetMessageId(mailbox: string, internetMessageId: string, folder?: string): Promise<MessageSummary[]> {
    // rfc822msgid: takes the id without angle brackets. Keep only Message-ID characters
    // (RFC 5322 atext, '.', '@') so a remote sender's value can't add search operators.
    const bare = internetMessageId.trim().replace(/^<|>$/g, "").replace(/[^A-Za-z0-9.@!#$%&'*+\/=?^_`{|}~-]/g, "");
    if (!bare) return [];
    return this.summariesFor(mailbox, `rfc822msgid:${bare}`, folder ? [GoogleEmailProvider.toLabelId(folder)] : undefined);
  }

  async sendConversationMessage(p: ConversationMessageParams): Promise<{ messageId: string }> {
    const gmail = this.getGmail(p.mailbox);
    const messageId = newMessageId(p.mailbox);
    const raw = buildMime({
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
    // Gmail files a reply into the sender's thread only when given its threadId.
    const parent = p.inReplyTo
      ? (await this.findByInternetMessageId(p.mailbox, p.inReplyTo).catch(() => [])).find(
          (m) => m.internetMessageId?.trim().toLowerCase() === p.inReplyTo!.trim().toLowerCase(),
        )
      : undefined;
    const sent = await gmail.users.messages.send({
      userId: "me",
      requestBody: { raw, threadId: parent?.conversationId },
    });
    // Report the Message-ID Gmail actually sent, in case it replaced ours.
    const meta = await gmail.users.messages
      .get({ userId: "me", id: sent.data.id!, format: "metadata", metadataHeaders: ["Message-ID"] })
      .catch(() => undefined);
    return { messageId: (meta && GoogleEmailProvider.getHeader(meta.data.payload?.headers ?? [], "Message-ID")) || messageId };
  }

  // A Gmail "folder" is a label; the mutually exclusive system ones act as locations.
  private static readonly LOCATION_LABELS = ["INBOX", "SPAM", "TRASH"];

  async moveMessage(mailbox: string, messageId: string, destinationFolder: string): Promise<{ id: string }> {
    const dest = GoogleEmailProvider.toLabelId(destinationFolder);
    const gmail = this.getGmail(mailbox);
    if (dest === "TRASH") {
      await gmail.users.messages.trash({ userId: "me", id: messageId });
      return { id: messageId };
    }
    await gmail.users.messages.modify({
      userId: "me",
      id: messageId,
      requestBody: {
        addLabelIds: [dest],
        removeLabelIds: GoogleEmailProvider.LOCATION_LABELS.filter((l) => l !== dest),
      },
    });
    return { id: messageId };
  }

  async getOrCreateFolder(mailbox: string, name: string, opts: { parentId?: string; hidden?: boolean } = {}): Promise<Folder> {
    const gmail = this.getGmail(mailbox);
    // Nested labels are named "Parent/Child".
    const parentName = opts.parentId
      ? (await gmail.users.labels.get({ userId: "me", id: GoogleEmailProvider.toLabelId(opts.parentId) })).data.name
      : undefined;
    const fullName = parentName ? `${parentName}/${name}` : name;
    const find = async () => {
      const labels = (await gmail.users.labels.list({ userId: "me" })).data.labels ?? [];
      const hit = labels.find((l) => l.name?.toLowerCase() === fullName.toLowerCase());
      if (!hit) return undefined;
      const d = (await gmail.users.labels.get({ userId: "me", id: hit.id! })).data;
      return { id: d.id!, name: d.name!, totalCount: d.messagesTotal ?? 0, unreadCount: d.messagesUnread ?? 0, childCount: 0 };
    };
    const existing = await find();
    if (existing) return existing;
    try {
      const created = await gmail.users.labels.create({
        userId: "me",
        requestBody: {
          name: fullName,
          labelListVisibility: opts.hidden ? "labelHide" : "labelShow",
          messageListVisibility: opts.hidden ? "hide" : "show",
        },
      });
      return { id: created.data.id!, name: created.data.name!, totalCount: 0, unreadCount: 0, childCount: 0 };
    } catch (err: any) {
      // Lost a create race (409: label name exists).
      if (httpStatus(err) === 409) {
        const raced = await find();
        if (raced) return raced;
      }
      throw err;
    }
  }
}

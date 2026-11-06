import { ConfidentialClientApplication } from "@azure/msal-node";
import { buildMimeMessage, newMessageId } from "../mime.js";
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

/**
 * One caller-supplied value (mailbox, message/attachment/folder id) as a single Graph path
 * segment: URL-encoded so it can't add segments, `?`/`#` or dot-segments — a crafted id such
 * as `../../other@x/messages/<id>` can't leave the mailbox the consumer authorized. `@` and `=`
 * stay literal (valid in a path segment, and what real mailboxes and Graph ids contain), so
 * ordinary paths are byte-for-byte what they were. Exported for tests; not public API.
 */
export function encodeSegment(value: string): string {
  if (typeof value !== "string" || value === "" || value === "." || value === "..") {
    throw new Error(`Invalid mailbox or id for a Graph path: ${JSON.stringify(value)}`);
  }
  return encodeURIComponent(value).replace(/%40/g, "@").replace(/%3D/g, "=");
}

/** Tagged template for Graph paths: every `${value}` goes through `encodeSegment`. */
export function gpath(strings: TemplateStringsArray, ...values: string[]): string {
  return strings.reduce((out, part, i) => out + encodeSegment(values[i - 1]) + part);
}

class GraphError extends Error {
  constructor(readonly status: number, readonly code: string | undefined, message: string, options?: ErrorOptions) {
    super(message, options);
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

  async sendEmail(params: SendParams) {
    const { mailbox, saveToSent } = params;
    await this.graph(gpath`/users/${mailbox}/sendMail`, {
      method: "POST",
      body: { message: this.newMessage(params), saveToSentItems: saveToSent ?? true },
    });

    return { from: mailbox };
  }

  /** Graph message JSON for a new message — shared by `sendEmail` and `createDraft`. */
  private newMessage({ to, subject, body, cc, bcc, attachments, origin }: DraftParams): Record<string, unknown> {
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
    return message;
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

    const r = await this.graph(gpath`/users/${mailbox}/mailFolders/${folderName}/messages`, { params });
    return {
      messages: (r.value as any[]).map(MicrosoftEmailProvider.toSummary),
      totalCount: r["@odata.count"] ?? r.value.length,
    };
  }

  /**
   * Keyset paging (see cursor.ts), not `@odata.nextLink`: Graph builds a message list's
   * nextLink from `$skip`, so a message dropping out mid-walk (read under unreadOnly,
   * moved) would shift the next one past the page boundary.
   */
  async listSince({ mailbox, since, folder, unreadOnly, limit, pageToken }: ListSinceParams): Promise<ListSinceResult> {
    const pageSize = pageSizeOf(limit);
    const prev = pageToken ? decodeCursor(pageToken, "microsoft") : undefined;
    const fromMs = prev ? prev.t : parseSince(since);
    const seen = new Set(prev?.ids ?? []);
    // Asking for seen + pageSize guarantees pageSize unseen messages whenever that many exist.
    const top = seen.size + pageSize;
    if (top > 1000) throw new Error("listSince: too many messages share one receive time to page past them");

    const folderName = MicrosoftEmailProvider.toFolderId(folder ?? "inbox");
    let filter = `receivedDateTime ge ${new Date(fromMs).toISOString()}`;
    if (unreadOnly) filter += " and isRead eq false";
    const r = await this.graph(gpath`/users/${mailbox}/mailFolders/${folderName}/messages`, {
      // Graph requires $orderby properties to lead the $filter, in the same order.
      params: { $filter: filter, $orderby: "receivedDateTime asc", $top: String(top), $select: SUMMARY_SELECT },
    });
    const value = r.value as any[];
    const page = value.filter((m) => !seen.has(m.id)).slice(0, pageSize).map(MicrosoftEmailProvider.toSummary);
    const more = value.length >= top || !!r["@odata.nextLink"];
    if (!more) return { messages: page };
    // A short server page holding only already-seen ids: no progress this call, but not
    // the end either — hand back the same cursor rather than claim there's nothing left.
    if (!page.length) return { messages: [], nextPageToken: encodeCursor({ p: "microsoft", t: fromMs, ids: [...seen] }) };
    const time = (m: MessageSummary) => Date.parse(m.receivedAt);
    return { messages: page, nextPageToken: encodeCursor(nextCursor("microsoft", page, time, (m) => m.id, prev)) };
  }

  async readMessage(mailbox: string, messageId: string): Promise<MessageFull> {
    const [r, attRes] = await Promise.all([
      this.graph(gpath`/users/${mailbox}/messages/${messageId}`, {
        params: {
          $select:
            "id,subject,from,sender,toRecipients,ccRecipients,bccRecipients,replyTo,internetMessageHeaders,conversationId,internetMessageId,parentFolderId,uniqueBody,receivedDateTime,isRead,body,hasAttachments",
          $expand: `singleValueExtendedProperties($filter=id eq ${odataString(MESSAGE_CLASS_PROP)})`,
        },
      }),
      this.graph(gpath`/users/${mailbox}/messages/${messageId}/attachments`, {
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
      gpath`/users/${mailbox}/messages/${messageId}/attachments/${attachmentId}`
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
      ? gpath`/users/${mailbox}/mailFolders/${MicrosoftEmailProvider.toFolderId(folder)}/messages`
      : gpath`/users/${mailbox}/messages`;

    const r = await this.graph(basePath, { params });
    return {
      messages: (r.value as any[]).map(MicrosoftEmailProvider.toSummary),
      totalCount: r["@odata.count"] ?? r.value.length,
    };
  }

  // An origin header can only be set when the message is created, and only the
  // create* (draft) actions were verified to take it — so origin routes via the draft.
  async replyToMessage(params: ReplyParams) {
    const { mailbox, messageId, body, replyAll, attachments, origin } = params;
    if (attachments?.length || origin !== undefined) {
      await this.sendDraft(mailbox, (await this.composeReplyDraft(params)).id);
      return;
    }
    const action = replyAll ? "replyAll" : "reply";
    await this.graph(gpath`/users/${mailbox}/messages/${messageId}/${action}`, {
      method: "POST",
      body: { comment: body },
    });
  }

  async forwardMessage(params: ForwardParams) {
    const { mailbox, messageId, to, comment, attachments, origin } = params;
    if (attachments?.length || origin !== undefined) {
      await this.sendDraft(mailbox, (await this.composeForwardDraft(params)).id);
      return;
    }
    await this.graph(gpath`/users/${mailbox}/messages/${messageId}/forward`, {
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
      const base = gpath`/users/${mailbox}/messages/${messageId}/attachments`;
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
        (await Promise.all(draftInlineIds.map((id) => this.graph(`${draftPath}/attachments/${encodeSegment(id)}`))))
          .map((a: any) => String(a?.contentId ?? "").replace(/^<|>$/g, "")),
      );
      const full = await Promise.all(inlineIds.map((id) => this.graph(`${base}/${encodeSegment(id)}`)));
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
   * take a `comment` string, so inline images can't ride on them. Instead the send path
   * builds the same draft `createReplyDraft`/`createForwardDraft` save (one mapping each,
   * below), then sends it. A failed send deletes the draft so no orphan is left behind
   * (the v0.7 behavior). Known hazard, tracked in #12: with `immutableIds`, if /send
   * errors after Exchange already accepted it (timeout, 5xx), that DELETE can delete the
   * message on its way out.
   */
  private async sendDraft(mailbox: string, draftId: string) {
    const draftPath = gpath`/users/${mailbox}/messages/${draftId}`;
    try {
      await this.graph(`${draftPath}/send`, { method: "POST" });
    } catch (err) {
      await this.graph(draftPath, { method: "DELETE" }).catch(() => {});
      throw err;
    }
  }

  private composeReplyDraft({ mailbox, messageId, body, replyAll, attachments, origin }: ReplyParams) {
    return this.composeDraft({
      mailbox, messageId, action: replyAll ? "createReplyAll" : "createReply", html: body, attachments: attachments ?? [], origin,
    });
  }

  private composeForwardDraft({ mailbox, messageId, to, comment, attachments, origin }: ForwardParams) {
    return this.composeDraft({
      mailbox, messageId, action: "createForward", html: comment ?? "", attachments: attachments ?? [], to, origin,
    });
  }

  /**
   * Create a reply/reply-all/forward draft (Graph fills recipients — Reply-To when set,
   * else From; reply-all adds To/Cc — threading, the quoted original and, on forward, the
   * original's attachments), prepend our HTML to its body, set forward recipients, add
   * the attachments. Never sends. A failure after the draft exists deletes it (it was never
   * sent). Returns the `create*` response overlaid with the PATCH response.
   */
  private async composeDraft({ mailbox, messageId, action, html, attachments, to, origin }: {
    mailbox: string;
    messageId: string;
    action: "createReply" | "createReplyAll" | "createForward";
    html: string;
    attachments: Attachment[];
    /** Forward recipients. */
    to?: string[];
    origin?: string;
  }): Promise<any> {
    const headers = this.originHeaders(origin);
    const draft = await this.graph(gpath`/users/${mailbox}/messages/${messageId}/${action}`, {
      method: "POST",
      body: "internetMessageHeaders" in headers ? { message: headers } : {},
    });
    // Nothing to clean up without an id; say so rather than blame the caller's input.
    if (typeof draft?.id !== "string" || !draft.id) throw new Error(`Graph returned no id for the ${action} draft`);
    const draftPath = gpath`/users/${mailbox}/messages/${draft.id}`;
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
      // PATCH returns the updated message (may carry internetMessageId the create lacked).
      const patched = await this.graph(draftPath, { method: "PATCH", body: patch });
      // createReply(All) drafts quote the original's HTML (incl. its cid: images) but don't
      // carry its inline attachments — copy them so the quote doesn't render broken images.
      // (createForward already carries every original attachment.)
      const carried = action === "createForward" ? [] : await this.originalInlineAttachments(mailbox, messageId, draftPath);
      // The original's quoted images are best-effort: a rejected upload just leaves that
      // quoted image broken. The caller's own attachments stay fail-hard.
      for (const att of carried) await this.addDraftAttachment(draftPath, att).catch(() => {});
      for (const att of attachments) await this.addDraftAttachment(draftPath, att);
      return { ...draft, ...(patched ?? {}) };
    } catch (err) {
      // Never sent (no /send on this path), so a plain DELETE is safe.
      await this.graph(draftPath, { method: "DELETE" }).catch(() => {});
      throw err;
    }
  }

  // ── Drafts ──
  // Built exactly as the matching send; nothing here calls /send or /sendMail.

  async createDraft(params: DraftParams): Promise<DraftRef> {
    // POST /messages creates the message in Drafts (isDraft: true) without sending.
    const draft = await this.graph(gpath`/users/${params.mailbox}/messages`, { method: "POST", body: this.newMessage(params) });
    return this.draftRef(params.mailbox, draft);
  }

  async createReplyDraft(params: ReplyParams): Promise<DraftRef> {
    return this.draftRef(params.mailbox, await this.composeReplyDraft(params));
  }

  async createForwardDraft(params: ForwardParams): Promise<DraftRef> {
    return this.draftRef(params.mailbox, await this.composeForwardDraft(params));
  }

  private async draftRef(mailbox: string, draft: any): Promise<DraftRef> {
    if (typeof draft?.id !== "string" || !draft.id) throw new Error("Graph returned no id for the created draft");
    let internetMessageId: string | undefined = draft?.internetMessageId ?? undefined;
    if (!internetMessageId) {
      // Not in the create response: read it back. Best-effort — the draft exists either way.
      const r = await this.graph(gpath`/users/${mailbox}/messages/${draft.id}`, { params: { $select: "internetMessageId" } }).catch(() => null);
      internetMessageId = r?.internetMessageId ?? undefined;
    }
    return { id: draft.id, draftId: draft.id, internetMessageId, conversationId: draft?.conversationId ?? undefined };
  }

  /**
   * Checks `isDraft` before deleting, so a received or sent message's id is refused. The
   * check and the DELETE are two separate requests, so this isn't atomic: see "Known race"
   * in the README.
   */
  async deleteDraft(mailbox: string, draftId: string): Promise<void> {
    if (!draftId.trim()) throw new Error("deleteDraft: draftId is empty");
    const path = gpath`/users/${mailbox}/messages/${draftId}`;
    const noDraft = (err: unknown): never => {
      // The item itself is gone (e.g. sent, so its id changed), as Gmail's deleteDraft reports.
      // Other 404s (ErrorInvalidUser, MailboxNotEnabledForRESTAPI…) say nothing about the draft.
      if (err instanceof GraphError && err.status === 404 && err.code === "ErrorItemNotFound") {
        throw new GraphError(404, err.code, `deleteDraft: no draft ${draftId} in ${mailbox} — already sent or discarded`, { cause: err });
      }
      // A malformed id (e.g. a Gmail draft id), as Gmail's deleteDraft reports.
      if (err instanceof GraphError && err.status === 400 && err.code === "ErrorInvalidIdMalformed") {
        throw new GraphError(400, err.code, `deleteDraft: no draft ${draftId} in ${mailbox} — not a draft id (pass DraftRef.draftId)`, { cause: err });
      }
      throw err;
    };
    const m = await this.graph(path, { params: { $select: "id,isDraft" } }).catch(noDraft);
    // status 404 ("no draft with this id"), the same shape Gmail's deleteDraft gives here.
    if (m?.isDraft !== true) throw new GraphError(404, "NotADraft", `deleteDraft: message ${draftId} is not a draft`);
    await this.graph(path, { method: "DELETE" }).catch(noDraft);
  }

  async markAsRead(mailbox: string, messageId: string, isRead: boolean) {
    await this.graph(gpath`/users/${mailbox}/messages/${messageId}`, {
      method: "PATCH",
      body: { isRead },
    });
  }

  async listFolders(mailbox: string, parentId?: string): Promise<Folder[]> {
    const path = parentId
      ? gpath`/users/${mailbox}/mailFolders/${parentId}/childFolders`
      : gpath`/users/${mailbox}/mailFolders`;

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
    const messageId = newMessageId(p.mailbox);
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
    await this.graph(gpath`/users/${p.mailbox}/sendMail`, {
      method: "POST",
      rawBody: Buffer.from(mime).toString("base64"),
      contentType: "text/plain",
    });
    return { messageId };
  }

  async findByInternetMessageId(mailbox: string, internetMessageId: string, folder?: string): Promise<MessageSummary[]> {
    const path = folder
      ? gpath`/users/${mailbox}/mailFolders/${MicrosoftEmailProvider.toFolderId(folder)}/messages`
      : gpath`/users/${mailbox}/messages`;
    const r = await this.graph(path, {
      params: { $filter: `internetMessageId eq ${odataString(internetMessageId)}`, $select: SUMMARY_SELECT, $top: "10" },
    });
    return (r.value as any[]).map(MicrosoftEmailProvider.toSummary);
  }

  async moveMessage(mailbox: string, messageId: string, destinationFolder: string): Promise<{ id: string }> {
    const r = await this.graph(gpath`/users/${mailbox}/messages/${messageId}/move`, {
      method: "POST",
      body: { destinationId: MicrosoftEmailProvider.toFolderId(destinationFolder) },
    });
    return { id: r.id };
  }

  async getOrCreateFolder(mailbox: string, name: string, opts: { parentId?: string; hidden?: boolean } = {}): Promise<Folder> {
    const base = opts.parentId
      ? gpath`/users/${mailbox}/mailFolders/${MicrosoftEmailProvider.toFolderId(opts.parentId)}/childFolders`
      : gpath`/users/${mailbox}/mailFolders`;
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

// ── Normalized types shared across all providers ──

export interface EmailAddress {
  address: string;
  name?: string;
}

export interface MessageSummary {
  id: string;
  subject: string;
  from: EmailAddress;
  to: EmailAddress[];
  cc: EmailAddress[];
  /** The provider's conversation/thread id (Graph `conversationId`, Gmail `threadId`).
   *  Groups a thread for display only — anyone who knows a subject and a Message-ID can
   *  join a thread, so never treat it as proof of membership. */
  conversationId?: string;
  /** RFC 5322 `Message-ID`, with angle brackets (e.g. `<abc@contoso.com>`). */
  internetMessageId?: string;
  /** Microsoft only: id of the folder holding the message. undefined on Google (labels). */
  parentFolderId?: string;
  receivedAt: string;
  isRead: boolean;
  preview: string;
  hasAttachments: boolean;
}

export interface AttachmentInfo {
  /** Provider-specific attachment ID. Use with `getAttachment` to download. */
  id: string;
  filename: string;
  contentType: string;
  size: number;
}

/** What sort of message this is, from headers and provider signals:
 *  - `ndr`: a non-delivery report / bounce;
 *  - `read-receipt`: a read or non-read notification;
 *  - `meeting`: a calendar invite, update, cancellation or response;
 *  - `auto-reply`: out-of-office or any other automatic message (`Auto-Submitted` other
 *    than `no`, `Precedence: auto_reply`, Exchange OOF/auto-reply class). Not
 *    `X-Auto-Response-Suppress`: Exchange adds it to ordinary distribution-group mail;
 *  - `message`: none of the above. */
export type MessageKind = "message" | "meeting" | "auto-reply" | "read-receipt" | "ndr";

export interface MessageFull {
  id: string;
  subject: string;
  from: EmailAddress;
  to: EmailAddress[];
  cc: EmailAddress[];
  /** See `MessageSummary.conversationId`. */
  conversationId?: string;
  /** See `MessageSummary.internetMessageId`. */
  internetMessageId?: string;
  /** See `MessageSummary.parentFolderId`. */
  parentFolderId?: string;
  /** Microsoft only: just what this message added to the thread (Graph `uniqueBody`),
   *  without quoted replies or forwarded content. Filled per its content type, like
   *  `bodyHtml`/`bodyText`. undefined on Google. */
  uniqueBodyHtml?: string;
  uniqueBodyText?: string;
  kind: MessageKind;
  /** Value of the provider config's `originHeader` on this message, trimmed. undefined
   *  when no `originHeader` is configured or the message doesn't carry it. NOT a security
   *  signal: anyone can put any X- header on mail they send. Use it to recognize your own
   *  sends (loops, self-echo), never to trust a sender. */
  origin?: string;
  bcc: EmailAddress[];
  /** `Reply-To` addresses — where a provider-native reply is actually sent (Graph
   *  honors it). Empty when the header is absent. Validate these, not just `from`. */
  replyTo: EmailAddress[];
  /** True when the message came through a mailing list / group (`List-Id` or `List-Post`
   *  header). There Reply-To usually points at the whole list, so callers composing a
   *  private reply may prefer From. */
  isMailingList: boolean;
  /** Microsoft 365 only: true iff the `from` address itself sent this message as an
   *  authenticated sender of the SAME organization. All of:
   *   - each header below is present exactly once with the expected value:
   *       X-MS-Exchange-Organization-AuthAs: Internal
   *       X-MS-Exchange-Organization-MessageDirectionality: Originating
   *       X-MS-Exchange-CrossTenant-FromEntityHeader: Hosted
   *   - the submitting mailbox (Graph `sender`) equals `from` — no send on behalf. (SendAs
   *     grants and app-only Mail.Send still yield sender = from: tenant-authorized, undetectable);
   *   - no sign it was resubmitted or auto-generated (rule redirect, auto-forward,
   *     `Resent-*`, `Auto-Submitted` other than `no`).
   *  Anything ambiguous → false. undefined = provider can't tell (Google).
   *  Unlike `from` alone (spoofable), Exchange strips these stamps from inbound mail. */
  internalSender?: boolean;
  receivedAt: string;
  isRead: boolean;
  bodyHtml: string;
  bodyText: string;
  hasAttachments: boolean;
  attachments: AttachmentInfo[];
}

export interface Folder {
  id: string;
  name: string;
  totalCount: number;
  unreadCount: number;
  childCount: number;
}

export interface ListResult {
  messages: MessageSummary[];
  totalCount: number;
}

/**
 * An attachment to be sent with an email.
 *
 * `content` is the raw bytes of the file as a `Buffer`. The library handles
 * any encoding required by the underlying provider (base64 for Gmail/Graph
 * wire formats) internally. Callers MUST NOT pre-encode.
 */
export interface Attachment {
  filename: string;
  content: Buffer;
  contentType: string;
  /**
   * Set to embed the file INLINE instead of attaching it: the HTML body
   * references it as `<img src="cid:<contentId>">` (e.g. a logo in an email
   * signature). Omit for a regular downloadable attachment. Bare id, no `<>`.
   */
  contentId?: string;
}

/**
 * Result of `getAttachment`. The `content` is raw bytes; callers can do
 * whatever they need with it (write to disk, forward through a channel,
 * base64-encode for transport, etc).
 *
 * Note: `filename` and `contentType` may be empty strings depending on
 * the provider — Gmail's attachment endpoint only returns `size` and
 * `content`. The caller usually already has these from the parent
 * `MessageFull.attachments[]` and should fall back to that.
 */
export interface RetrievedAttachment {
  filename: string;
  contentType: string;
  size: number;
  content: Buffer;
}

/** Stamped as the provider config's `originHeader` (required when this is set) when the
 *  message is created. 1–200 printable ASCII characters, e.g. `agentfleet; kind=digest`. */
type Origin = { origin?: string };

export interface SendParams extends Origin {
  mailbox: string;
  to: string[];
  subject: string;
  body: string;
  cc?: string[];
  bcc?: string[];
  saveToSent?: boolean;
  attachments?: Attachment[];
}

/**
 * A message sent as raw MIME with caller-controlled threading: no quoted original, no
 * recipients derived from the parent (never from `Reply-To`), our own `Message-ID`.
 * Omit `inReplyTo`/`references` to start a new thread.
 */
export interface ConversationMessageParams extends Origin {
  mailbox: string;
  to: string[];
  cc?: string[];
  subject: string;
  /** HTML body. */
  html: string;
  /** Plain-text alternative; sent as `multipart/alternative` with the HTML when set. */
  text?: string;
  /** Parent's `Message-ID`, with angle brackets. */
  inReplyTo?: string;
  /** Space-separated `Message-ID`s of the thread. Defaults to `inReplyTo` when that is set. */
  references?: string;
  /** Adds `Auto-Submitted: <value>` (RFC 3834) and `X-Auto-Response-Suppress: All`, so
   *  auto-responders (out-of-office, other bots) don't answer it. */
  autoSubmitted?: "auto-replied" | "auto-generated";
  /** Inline (`contentId`) images, e.g. a signature, and regular attachments. */
  attachments?: Attachment[];
}

export interface ListParams {
  mailbox: string;
  folder?: string;
  limit?: number;
  offset?: number;
  unreadOnly?: boolean;
}

export interface SearchParams {
  mailbox: string;
  query: string;
  folder?: string;
  limit?: number;
  offset?: number;
}

export interface ReplyParams extends Origin {
  mailbox: string;
  messageId: string;
  body: string;
  replyAll?: boolean;
  /** Extra attachments for the reply — typically inline (`contentId`) images such
   *  as a signature logo referenced from `body`. */
  attachments?: Attachment[];
}

export interface ForwardParams extends Origin {
  mailbox: string;
  messageId: string;
  to: string[];
  comment?: string;
  /** Extra attachments added to the forward (on top of the original's, which are
   *  always carried) — typically inline (`contentId`) images referenced from `comment`. */
  attachments?: Attachment[];
}

// ── Provider configuration ──

interface CommonProviderConfig {
  /** Custom header (`X-…`) that `origin` on a send is stamped as and that
   *  `MessageFull.origin` is read from, e.g. `"X-AgentFleet-Origin"`. */
  originHeader?: string;
}

export interface GoogleProviderConfig extends CommonProviderConfig {
  type: "google";
  /**
   * Parsed Google service account key (the JSON object, not a string).
   * The service account must have domain-wide delegation enabled and the
   * `https://www.googleapis.com/auth/gmail.modify` and `gmail.send` scopes
   * authorized in the Google Workspace admin console.
   */
  serviceAccountKey: Record<string, unknown>;
}

export interface MicrosoftProviderConfig extends CommonProviderConfig {
  type: "microsoft";
  /** App registration client ID. */
  clientId: string;
  /** App registration client secret. */
  clientSecret: string;
  /** Azure tenant ID. */
  tenantId: string;
  /**
   * Send `Prefer: IdType="ImmutableId"` on every Graph call: message ids then survive a
   * move between folders. Changes the format of every id the provider returns, so turn it
   * on before storing ids, not midway.
   */
  immutableIds?: boolean;
}

export type ProviderConfig = GoogleProviderConfig | MicrosoftProviderConfig;

// ── Provider contract ──
// Every method receives the target mailbox explicitly.
// For Microsoft: the MSAL client has tenant-wide access via app permissions.
// For Google: a Gmail client is created per delegated user (cached).

export interface EmailProvider {
  readonly name: string;

  sendEmail(params: SendParams): Promise<{ from: string }>;
  listMessages(params: ListParams): Promise<ListResult>;
  readMessage(mailbox: string, messageId: string): Promise<MessageFull>;
  searchMessages(params: SearchParams): Promise<ListResult>;
  replyToMessage(params: ReplyParams): Promise<void>;
  forwardMessage(params: ForwardParams): Promise<void>;
  markAsRead(mailbox: string, messageId: string, isRead: boolean): Promise<void>;
  listFolders(mailbox: string, parentId?: string): Promise<Folder[]>;
  getAttachment(mailbox: string, messageId: string, attachmentId: string): Promise<RetrievedAttachment>;
  /** Send raw MIME with caller-controlled threading. Returns the sent `Message-ID`. */
  sendConversationMessage(params: ConversationMessageParams): Promise<{ messageId: string }>;
  /** Messages whose `Message-ID` is `internetMessageId`, optionally within one folder.
   *  A just-sent message's Sent copy can take seconds to appear — retry on []. */
  findByInternetMessageId(mailbox: string, internetMessageId: string, folder?: string): Promise<MessageSummary[]>;
  /** Move a message to a folder (id or neutral name). Returns its id afterwards, which on
   *  Microsoft changes unless `immutableIds` is on. */
  moveMessage(mailbox: string, messageId: string, destinationFolder: string): Promise<{ id: string }>;
  /** The folder named `name` (under `parentId` if given), created if missing. `hidden`
   *  applies only on creation (Microsoft `isHidden`; Gmail hides the label). */
  getOrCreateFolder(mailbox: string, name: string, opts?: { parentId?: string; hidden?: boolean }): Promise<Folder>;
}

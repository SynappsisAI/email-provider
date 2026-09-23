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

export interface MessageFull {
  id: string;
  subject: string;
  from: EmailAddress;
  to: EmailAddress[];
  cc: EmailAddress[];
  bcc: EmailAddress[];
  /** `Reply-To` addresses — where a provider-native reply is actually sent (Graph
   *  honors it). Empty when the header is absent. Validate these, not just `from`. */
  replyTo: EmailAddress[];
  /** True when the message came through a mailing list / group (`List-Id` or `List-Post`
   *  header). There Reply-To usually points at the whole list, so callers composing a
   *  private reply may prefer From. */
  isMailingList: boolean;
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

export interface SendParams {
  mailbox: string;
  to: string[];
  subject: string;
  body: string;
  cc?: string[];
  bcc?: string[];
  saveToSent?: boolean;
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

export interface ReplyParams {
  mailbox: string;
  messageId: string;
  body: string;
  replyAll?: boolean;
  /** Extra attachments for the reply — typically inline (`contentId`) images such
   *  as a signature logo referenced from `body`. */
  attachments?: Attachment[];
}

export interface ForwardParams {
  mailbox: string;
  messageId: string;
  to: string[];
  comment?: string;
  /** Extra attachments added to the forward (on top of the original's, which are
   *  always carried) — typically inline (`contentId`) images referenced from `comment`. */
  attachments?: Attachment[];
}

// ── Provider configuration ──

export interface GoogleProviderConfig {
  type: "google";
  /**
   * Parsed Google service account key (the JSON object, not a string).
   * The service account must have domain-wide delegation enabled and the
   * `https://www.googleapis.com/auth/gmail.modify` and `gmail.send` scopes
   * authorized in the Google Workspace admin console.
   */
  serviceAccountKey: Record<string, unknown>;
}

export interface MicrosoftProviderConfig {
  type: "microsoft";
  /** App registration client ID. */
  clientId: string;
  /** App registration client secret. */
  clientSecret: string;
  /** Azure tenant ID. */
  tenantId: string;
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
}

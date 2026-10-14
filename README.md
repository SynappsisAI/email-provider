# @synappsis/email-provider

Provider-agnostic email library for Synappsis services. Wraps Google Workspace
(Gmail) and Microsoft 365 (Graph) behind a single `EmailProvider` interface.

This is a **pure mechanics** library:

- ✅ Send, list, read, search, reply, forward, get attachments, mark read, list folders
- ✅ Per-mailbox impersonation via Google domain-wide delegation
- ✅ Tenant-wide app permissions via Microsoft Graph
- ✅ MIME construction with attachments
- ✅ Buffer-based attachments (no base64 in the public API)
- ❌ No HTTP server, no MCP transport, no Express
- ❌ No authorization / policy / API key validation
- ❌ No rate limiting per caller (retries with backoff for provider 429s only)

Authorization, transport, and policy belong to the consumer (e.g. synmail's
HTTP server with API key whitelists, or AgentFleet's per-tenant channel config).

---

## Install

The package is distributed as a git URL dependency. Pin to a tag for
reproducibility:

```jsonc
// package.json
{
  "dependencies": {
    "@synappsis/email-provider": "github:SynappsisAI/email-provider#v0.1.0"
  }
}
```

For local development against a checkout:

```jsonc
{
  "dependencies": {
    "@synappsis/email-provider": "file:../email-provider"
  }
}
```

## Usage

```typescript
import { createEmailProvider, type Attachment } from "@synappsis/email-provider";
import { readFileSync } from "node:fs";

// Construct from a typed config. The consumer sources credentials however
// it wants — env vars, AWS Secrets Manager, a database, etc.
const provider = createEmailProvider({
  type: "google",
  serviceAccountKey: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY!),
});

// Or Microsoft:
// const provider = createEmailProvider({
//   type: "microsoft",
//   clientId: process.env.MS_APP_ID!,
//   clientSecret: process.env.MS_CLIENT_SECRET!,
//   tenantId: process.env.MS_TENANT_ID!,
// });

// Send an email with a real binary attachment — no base64 in sight.
const pdfBytes: Buffer = readFileSync("/tmp/report.pdf");

await provider.sendEmail({
  mailbox: "jarvis@synappsis.ai",
  to: ["rafa@synappsis.ai"],
  subject: "Pipeline report",
  body: "<p>Here's the report you asked for.</p>",
  attachments: [{
    filename: "report.pdf",
    contentType: "application/pdf",
    content: pdfBytes,  // Buffer, not base64
  }],
});

// Read an inbox.
const inbox = await provider.listMessages({
  mailbox: "jarvis@synappsis.ai",
  folder: "inbox",
  limit: 10,
  unreadOnly: true,
});

// Download an attachment — get raw bytes back.
const message = await provider.readMessage("jarvis@synappsis.ai", inbox.messages[0].id);
const attached = await provider.getAttachment(
  "jarvis@synappsis.ai",
  message.id,
  message.attachments[0].id,
);
// attached.content is a Buffer.
```

## The contract

```typescript
interface EmailProvider {
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
  sendConversationMessage(params: ConversationMessageParams): Promise<{ messageId: string }>;
  listSince(params: ListSinceParams): Promise<ListSinceResult>;
  findByInternetMessageId(mailbox: string, internetMessageId: string, folder?: string): Promise<MessageSummary[]>;
  moveMessage(mailbox: string, messageId: string, destinationFolder: string): Promise<{ id: string }>;
  getOrCreateFolder(mailbox: string, name: string, opts?: { parentId?: string; hidden?: boolean }): Promise<Folder>;
}
```

Folder arguments accept a real folder/label id or a neutral name: `inbox`, `sent`,
`drafts`/`draft`, `trash`, `spam`/`junk` (plus `starred`/`important` on Gmail).

`Attachment` and `RetrievedAttachment` use `Buffer` for the file bytes.
**Callers never see base64.** The library handles whatever encoding the
underlying provider needs internally:

- **Google**: base64-encodes the buffer inside the MIME multipart body
  before passing to `gmail.users.messages.send`
- **Microsoft**: base64-encodes the buffer for Graph API's `contentBytes`
  field on `#microsoft.graph.fileAttachment`

### Inline images (CID)

Set `contentId` on an attachment to embed it in the HTML body instead of
attaching it — e.g. a logo in a signature. Reference it as `cid:<contentId>`:

```ts
await provider.sendEmail({
  mailbox: "alfred@matelpa.com",
  to: ["cliente@example.com"],
  subject: "Hola",
  body: '<p>Saludos,</p><img src="cid:logo" alt="Logo">',
  attachments: [{ filename: "logo.png", contentType: "image/png", content: logoBytes, contentId: "logo" }],
});
```

- **Google**: the HTML and its inline parts go in a `multipart/related` (with
  `Content-ID` + `Content-Disposition: inline`), wrapped in `multipart/mixed`
  when there are also regular attachments.
- **Microsoft**: `isInline: true` + `contentId` on the `fileAttachment`.

`replyToMessage` / `forwardMessage` also take `attachments` (e.g. a signed reply). On
Microsoft this switches to Graph's draft flow (`createReply`/`createForward` → patch body →
add attachments → send) since the one-shot `/reply` and `/forward` only take a `comment`.
Each attachment on that path is a single Graph upload, so it must be **< 3MB** (fine for a
signature logo; larger files would need an upload session, not implemented). Reply drafts
also get the original's inline images copied in, so the quoted part renders. On Gmail, a
forward carries the original's attachments while the encoded message fits in ~24MB; above
that it goes out body-only (the pre-0.4.0 behavior). `MessageFull.replyTo` exposes the `Reply-To` header —
where a native reply is actually delivered (both providers reply to Reply-To when set,
else From) — so callers can validate it.

### Threads, ids and message kinds

`MessageSummary` and `MessageFull` carry `cc`, `conversationId` (Graph `conversationId`,
Gmail `threadId`), `internetMessageId` (the `Message-ID`, with `<>`) and, on Microsoft,
`parentFolderId`. Never treat `conversationId` as proof that a message belongs to a
thread: anyone who knows a subject and a `Message-ID` can join one.

`MessageFull` adds:

- `uniqueBodyHtml` / `uniqueBodyText` (Microsoft only): just what this message added,
  without quoted replies or forwarded content (Graph `uniqueBody`).
- `kind`: `message`, `meeting`, `auto-reply` (out-of-office, `Auto-Submitted` other than
  `no`, `Precedence: auto_reply`), `read-receipt` or `ndr`.
  Microsoft also uses Exchange's message class (`IPM.Schedule.Meeting.*`,
  `IPM.Note.Rules.OofTemplate.*`, `REPORT.*.IPNRN`, `REPORT.*.NDR`). Gmail spots meetings
  by an invite's `text/calendar` part; an attached `.ics` file doesn't count.
  `X-Auto-Response-Suppress` is deliberately ignored: Exchange adds it to ordinary mail
  delivered through distribution groups, so it would drop real messages.

**Immutable ids (Microsoft).** Set `immutableIds: true` in the Microsoft config to send
`Prefer: IdType="ImmutableId"` on every Graph call: a message id then survives a move to
another folder. It changes the format of every id the provider returns, so turn it on
before you store ids, not midway. Without it, `moveMessage` returns the new id.

**Search and folders (Microsoft).** Graph can't combine `$search` with a `parentFolderId`
filter. To keep a folder out of search results, filter on `parentFolderId` client-side.

### Polling for new mail: `listSince`

To pick up everything that arrived after a watermark, use `listSince`, not
`listMessages` with `offset`:

```ts
let pageToken: string | undefined;
do {
  const page = await provider.listSince({ mailbox, since: watermark, unreadOnly: true, limit: 50, pageToken });
  for (const m of page.messages) handle(m);   // oldest first
  pageToken = page.nextPageToken;             // absent = caught up
} while (pageToken /* and this tick's batch budget allows */);
```

- `since` is inclusive (`receivedAt >= since`). Messages sharing the watermark's timestamp
  come back again on the next walk, so de-dupe them by id.
- Pages are ascending by `receivedAt` and continue exactly where the previous one stopped,
  even if messages are marked read, moved or deleted between calls. A page can be short or
  empty and still carry a `nextPageToken`; only a missing token means the end.
- The token is a keyset cursor: the time of the last message returned plus the ids already
  returned at that time. Offset paging can't be made correct here. Graph's message
  `@odata.nextLink` is `$skip`-based, so a message dropping out mid-walk shifts the next one
  past the page boundary. Gmail caps `maxResults` at 500 and doesn't guarantee list order.
- On Gmail, `after:`/`before:` only have second granularity and list results carry no dates.
  Each page finds a time window with few enough ids, fetches their metadata (at most 10
  `messages.get` in flight), and filters `internalDate` exactly.
- A message moved INTO the folder keeps its original receive time, so if that time is older
  than the watermark, the walk won't see it.
- Every watermark poller trusts that mail shows up in receive-time order. If a message with
  an earlier timestamp could become listable after a later one was already returned (e.g.
  search-index lag), overlap the watermark by a few minutes and de-dupe by id.

### Origin header

Set `originHeader` in either provider's config (an `X-` header, e.g.
`"X-AgentFleet-Origin"`) to mark your own sends. Pass `origin` (1–200 printable ASCII
characters, e.g. `agentfleet; kind=digest`) to `sendEmail`, `replyToMessage`,
`forwardMessage` or `sendConversationMessage`, and the message carries
`<originHeader>: <origin>`. Passing `origin` without `originHeader` configured throws.
`MessageFull.origin` reads it back.

The header is set when the message is created: Graph rejects adding it later. On
Microsoft, `sendEmail` puts it in `internetMessageHeaders`. A reply or forward with
`origin` goes through `createReply`/`createReplyAll`/`createForward` with the header in
the request, then sends that draft, even without attachments. Graph only accepts
custom headers named `X-…`, which is why `originHeader` must start with `X-`.

`origin` is **not** a security signal. Anyone can put any X- header on mail they send.
Use it to recognize your own messages (loops, self-echo), never to trust a sender. For
that, see `internalSender`.

### Conversation messages (raw MIME)

`sendConversationMessage` sends a message as raw MIME (Graph MIME `sendMail`, Gmail
`messages.send`) with threading you control:

- `inReplyTo` / `references` thread it into an existing conversation. Omit them to start a
  new one.
- recipients are exactly `to` / `cc`. Nothing is derived from the parent or its `Reply-To`.
- there's no quoted original. Graph and Gmail keep the generated `Message-ID`, which is
  returned so you can find the Sent copy later with `findByInternetMessageId`. On Microsoft
  the copy shows up in Sent Items a few seconds after sending, so retry until it does.
- `text` adds a plain-text alternative. Inline (`contentId`) and regular attachments work
  as in `sendEmail`.
- `autoSubmitted: "auto-replied"` adds `Auto-Submitted` and `X-Auto-Response-Suppress: All`,
  so out-of-office replies and other bots don't answer it. Such a message then reads back
  with `kind: "auto-reply"` and `internalSender: false`. Recognize your own sends by
  `origin`.

`getOrCreateFolder` finds a folder by name, including hidden ones, or creates it.
`hidden: true` only takes effect at creation. On Microsoft it sets `isHidden`. On Gmail it
hides the label, and `parentId` nests it as `Parent/Name`. `moveMessage` on Gmail adds the
destination label and removes `INBOX`/`SPAM`/`TRASH`.

### Internal senders (Microsoft 365)

`From` is unauthenticated and spoofable, and DKIM/DMARC can't tell internal mail apart
(Exchange doesn't sign mail that never leaves the organization: `dkim=none`, `dmarc=none`).
On Microsoft, `readMessage` sets `MessageFull.internalSender` to `true` only when the `From`
address itself sent the message as an authenticated sender of the **same** organization.
First, Exchange must have stamped it — each of these headers present exactly once with this
value (case-insensitive):

| Header | Internal | External |
|---|---|---|
| `X-MS-Exchange-Organization-AuthAs` | `Internal` | `Anonymous` |
| `X-MS-Exchange-Organization-MessageDirectionality` | `Originating` | `Incoming` |
| `X-MS-Exchange-CrossTenant-FromEntityHeader` | `Hosted` | `Internet` |

Exchange strips or overwrites these headers on inbound mail, so they can't be forged from
outside. But they describe whoever **submitted** the message, not the owner of `From`, so
it must also hold that:

- Graph's `sender` equals `from` (rules out send-on-behalf), and
- nothing shows the message was resubmitted or auto-generated: no
  `X-MS-Exchange-Organization-AutoForwarded`, `X-MS-Exchange-Inbox-Rules-Loop`,
  `X-MS-Exchange-Generated-Message-Source` or `Resent-*` header, and `Auto-Submitted`
  absent or `no`. (An inbox-rule redirect of a spoofed outside message would otherwise be
  resubmitted as internal, keeping the spoofed `From`.)

Not detectable: a mailbox with **SendAs** rights on another, or an app with app-only
`Mail.Send`, submits as that address and Graph reports `sender = from`. That impersonation
is granted by the tenant admin; consumers that care must restrict who holds those grants.

Anything missing, duplicated, contradictory or ambiguous → `false`. On Google it is `undefined`
(no equivalent signal wired yet). Raw headers are deliberately not exposed.

## Provider setup

### Google Workspace

Requires a service account with **domain-wide delegation** enabled and the
following OAuth scopes authorized in the Google Workspace admin console:

- `https://www.googleapis.com/auth/gmail.modify`
- `https://www.googleapis.com/auth/gmail.send`

Pass the parsed key JSON via `serviceAccountKey`.

### Microsoft 365

Requires an Azure AD app registration with **application** (not delegated)
permissions:

- `Mail.Send`
- `Mail.ReadWrite`

Grant admin consent, then pass `clientId`, `clientSecret`, and `tenantId`.

## Distribution / build

There is no build step. The package's `main` and `types` both point to
`src/index.ts`. Consumers import the TypeScript source directly:

- **synmail** runs the source via `tsx`
- **AgentFleet** bundles the source via esbuild (CDK NodejsFunction)

This avoids a `dist/` step and keeps git URL installs simple.

## Versioning

Releases are git tags (`v0.1.0`, `v0.2.0`, etc). Bump `package.json` to
match the tag in the same commit. Consumers pin to a specific tag in
their `dependencies` field.

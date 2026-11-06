# @synappsis/email-provider

Provider-agnostic email library for Synappsis services. Wraps Google Workspace
(Gmail) and Microsoft 365 (Graph) behind a single `EmailProvider` interface.

This is a **pure mechanics** library:

- ✅ Send, list, read, search, reply, forward, get attachments, mark read, list folders
- ✅ Drafts: create (new / reply / reply-all / forward) without sending, delete, send
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
  createDraft(params: DraftParams): Promise<DraftRef>;
  createReplyDraft(params: ReplyParams): Promise<DraftRef>;
  createForwardDraft(params: ForwardParams): Promise<DraftRef>;
  deleteDraft(mailbox: string, draftId: string): Promise<void>;
  sendDraft(mailbox: string, draftId: string): Promise<void>;
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
also get the original's inline images copied in, so the quoted part renders. If the final
`/send` fails, the draft is deleted so no orphan is left behind. On Gmail, a
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

### Drafts

Leave a message in a mailbox's Drafts for its owner to review and send themselves:

```ts
const ref = await provider.createReplyDraft({ mailbox, messageId, body: "<p>Propuesta…</p>", replyAll: true });
// ref = { id, draftId, internetMessageId?, conversationId? }
const check = await provider.readMessage(mailbox, ref.id);   // to / cc / bcc / replyTo as saved
await provider.deleteDraft(mailbox, ref.draftId);             // e.g. to replace it
```

- `createDraft` takes `sendEmail`'s params (minus `saveToSent`), `createReplyDraft` takes
  `replyToMessage`'s and `createForwardDraft` takes `forwardMessage`'s. Each draft is built
  as that send would build it: the same recipients (a reply goes to the original's
  `Reply-To` when set, else its `From`; reply-all adds the original's To/Cc minus the
  mailbox), threading, quoting and attachments. On Gmail a reply doesn't quote the
  original, the same as `replyToMessage`.
- The `create*` methods **never send**, on any path. Nothing in them calls a send endpoint. On Microsoft
  a failure after the draft exists tries to delete it, so it isn't left half-built; if
  that delete also fails, the draft can remain.
- No origin header is added unless you pass `origin`.
- `DraftRef.id` is the draft's message id (for `readMessage`). `draftId` is what
  `deleteDraft` and `sendDraft` take: on Gmail the drafts-API id, which stays put when the draft is edited
  (the message id doesn't); on Microsoft the message id again. Always pass `draftId`:
  `id` happens to work on Microsoft but is refused on Gmail. `internetMessageId` is the
  draft's `Message-ID` as created: on Gmail one we generate (always set), on Microsoft
  Graph's (undefined if Graph didn't report one). It isn't guaranteed to survive the owner
  editing or sending the draft (Gmail replaces a draft's message on every edit), so don't
  rely on it alone to tell whether a draft was sent.
- `deleteDraft` refuses anything that isn't a draft when it checks. Microsoft checks
  `isDraft` first and throws on anything else. Gmail goes through `drafts.delete`, which only
  knows draft ids, so a message id is refused.
- Treat a deleted draft as not restorable. Gmail deletes it permanently. On Microsoft, a live
  run on 2026-11-05 ([#10](https://github.com/SynappsisAI/email-provider/issues/10)) found the
  deleted draft in neither Drafts nor Deleted Items, while `GET` by its immutable id still
  answered 200. So with `immutableIds`, `readMessage` can't confirm a draft is gone, and a
  second `deleteDraft` on the same id is untested.
- **Known race (Microsoft).** The `isDraft` check and the `DELETE` are
  two requests. In that same run, with `immutableIds`, a `DELETE` with a stale `If-Match` still
  deleted the message, so don't rely on `If-Match` to tie them together. With `immutableIds`, a
  draft the person sends in that instant keeps its id, and the sent message is deleted. A draft
  edited in that instant is deleted too, and whether a send still queued reads as a draft
  wasn't tested. Without it the id should change when
  the message moves to Sent Items, so a late `DELETE` would get a 404 and `deleteDraft` would
  report "no draft". That wasn't tested live, and Exchange submits asynchronously, so don't
  treat it as race-free.
- Microsoft: a new-message draft is `POST /messages` (attachments ride in the same request,
  as in `sendEmail`). Reply/forward drafts are `createReply`/`createReplyAll`/
  `createForward` + body patch + attachments, the flow described above minus the `/send`,
  so each added attachment must be < 3MB. (Without attachments or `origin`,
  `replyToMessage`/`forwardMessage` use Graph's one-shot `/reply`/`/forward`, where Graph
  places the comment above the quote itself; the result is equivalent, not byte-identical.)

#### Sending a draft

`sendDraft` sends an existing draft as it's stored, without composing anything new, so
threading and the recipients the provider resolved stay as they are. It's what a reply needs
when those recipients must be checked first:

```ts
const ref = await provider.createReplyDraft({ mailbox, messageId, body, replyAll: true });
const draft = await provider.readMessage(mailbox, ref.id);   // recipients the provider resolved
if (allowed(draft.to, draft.cc, draft.bcc)) await provider.sendDraft(mailbox, ref.draftId);
else await provider.deleteDraft(mailbox, ref.draftId);
```

- It takes `DraftRef.draftId` on both providers, as `deleteDraft` does.
- The draft keeps the `origin` it was created with and goes to Sent as a normal send.
- It refuses anything that isn't a draft, with the same errors as `deleteDraft`. Microsoft
  checks `isDraft` first (`POST /messages/{id}/send` then sends it). Gmail uses `drafts.send`,
  which only knows draft ids.
- **A failed send never deletes the draft.** The caller owns it, and a timeout or 5xx may
  still have gone out, so check before retrying: Sent, and on Microsoft the Outbox too, since
  Exchange submits asynchronously and a message on its way out may not be in Sent yet. Whether
  such a message still reads as a draft wasn't tested, so a retry in that window could send it
  twice. (`replyToMessage`/`forwardMessage` are different: they created their draft themselves
  and delete it if `/send` fails.)
- **What you read isn't locked.** An edit made between your read and the send goes out as
  edited, including added recipients, so this narrows the gap but isn't a hard allowlist.
  There's no atomic version check: Gmail's `drafts.send` takes none, and Graph ignored
  `If-Match` on message `DELETE` (see above), so it isn't relied on here. On Microsoft the
  `isDraft` check and the `/send` are also two requests. Keep the gap between read and send
  short.

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

`gmail.modify` also covers the drafts API (`drafts.create` / `drafts.delete` / `drafts.send`), so drafts
need no extra scope. (`gmail.compose` would cover it too, but isn't needed with `modify`.)

Pass the parsed key JSON via `serviceAccountKey`.

### Microsoft 365

Requires an Azure AD app registration with **application** (not delegated)
permissions:

- `Mail.Send`
- `Mail.ReadWrite` (also what creating and deleting drafts needs)

`sendDraft` needs both: it reads the message to check `isDraft`, then sends it.

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

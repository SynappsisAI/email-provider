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
}
```

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

import { test, beforeEach, afterEach } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { MicrosoftEmailProvider } from "../src/providers/microsoft.js";

// Synthetic, anonymized Graph message payloads (shaped like readMessage's response).
const fx = JSON.parse(readFileSync(new URL("./fixtures/graph-messages.json", import.meta.url), "utf8"));

type Call = { method: string; url: URL; headers: Record<string, string>; body: string | undefined };
let calls: Call[] = [];
let responses: ((c: Call) => { status?: number; json?: unknown } | undefined)[] = [];
const realFetch = globalThis.fetch;

beforeEach(() => {
  calls = [];
  responses = [];
  globalThis.fetch = (async (input: string, init: RequestInit = {}) => {
    const call: Call = {
      method: init.method ?? "GET",
      url: new URL(input),
      headers: init.headers as Record<string, string>,
      body: init.body as string | undefined,
    };
    calls.push(call);
    const r = responses.map((f) => f(call)).find(Boolean) ?? {};
    const status = r.status ?? 200;
    const text = r.json === undefined ? "" : JSON.stringify(r.json);
    return new Response(status === 204 ? null : text, { status: status === 204 ? 204 : status });
  }) as typeof fetch;
});
afterEach(() => {
  globalThis.fetch = realFetch;
});

const on = (method: string, pathRe: RegExp, res: { status?: number; json?: unknown }) =>
  responses.push((c) => (c.method === method && pathRe.test(c.url.pathname) ? res : undefined));

function provider(extra: { immutableIds?: boolean; originHeader?: string } = {}) {
  const p = new MicrosoftEmailProvider({ type: "microsoft", clientId: "c", clientSecret: "s", tenantId: "t", ...extra });
  (p as any).getToken = async () => "token";
  return p;
}

const toFull = (m: unknown, originHeader?: string) => (MicrosoftEmailProvider as any).toFull(m, originHeader);

// ── Normalization ──

test("toFull: thread fields, uniqueBody, cc, origin, kind", () => {
  const m = toFull(fx.reply, "X-AgentFleet-Origin");
  assert.equal(m.conversationId, fx.reply.conversationId);
  assert.equal(m.internetMessageId, fx.reply.internetMessageId);
  assert.equal(m.parentFolderId, fx.reply.parentFolderId);
  assert.equal(m.uniqueBodyHtml, "<html><body><p>Sí, confirmado.</p></body></html>");
  assert.equal(m.uniqueBodyText, "");
  assert.deepEqual(m.cc, [{ address: "carol@contoso.com", name: "Carol" }]);
  assert.equal(m.origin, "agentfleet; kind=reply");
  assert.equal(m.kind, "message");
});

test("toFull: origin undefined without originHeader configured", () => {
  assert.equal(toFull(fx.reply).origin, undefined);
});

test("toFull: kind for meeting, OOF, read receipt, NDR (message class as Graph echoes it)", () => {
  assert.equal(toFull(fx.meeting).kind, "meeting");
  assert.equal(toFull(fx.oof).kind, "auto-reply");
  assert.equal(toFull(fx["read-receipt"]).kind, "read-receipt");
  assert.equal(toFull(fx.ndr).kind, "ndr");
});

test("listMessages: summaries carry cc and thread fields", async () => {
  on("GET", /\/mailFolders\/inbox\/messages$/, { json: { value: [fx.reply], "@odata.count": 1 } });
  const r = await provider().listMessages({ mailbox: "agent@contoso.com" });
  assert.deepEqual(r.messages[0].cc, [{ address: "carol@contoso.com", name: "Carol" }]);
  assert.equal(r.messages[0].conversationId, fx.reply.conversationId);
  assert.equal(r.messages[0].internetMessageId, fx.reply.internetMessageId);
  assert.equal(r.messages[0].parentFolderId, fx.reply.parentFolderId);
  assert.match(calls[0].url.searchParams.get("$select")!, /ccRecipients.*conversationId.*internetMessageId.*parentFolderId/);
});

test("readMessage selects uniqueBody and expands the message class", async () => {
  on("GET", /\/messages\/AAMkReply==$/, { json: fx.reply });
  on("GET", /\/attachments$/, { json: { value: [] } });
  const m = await provider().readMessage("agent@contoso.com", "AAMkReply==");
  assert.equal(m.kind, "message");
  const get = calls.find((c) => c.url.pathname.endsWith("AAMkReply=="))!;
  assert.match(get.url.searchParams.get("$select")!, /uniqueBody/);
  assert.equal(get.url.searchParams.get("$expand"), "singleValueExtendedProperties($filter=id eq 'String 0x001A')");
});

// ── Immutable ids ──

test("immutableIds sends Prefer: IdType=ImmutableId on every call; off by default", async () => {
  on("GET", /./, { json: { value: [] } });
  await provider({ immutableIds: true }).listFolders("agent@contoso.com");
  await provider().listFolders("agent@contoso.com");
  assert.equal(calls[0].headers.Prefer, 'IdType="ImmutableId"');
  assert.equal(calls[1].headers.Prefer, undefined);
});

// ── Origin on existing send paths ──

test("sendEmail: origin goes in message.internetMessageHeaders", async () => {
  on("POST", /\/sendMail$/, { status: 202 });
  await provider({ originHeader: "X-AgentFleet-Origin" }).sendEmail({
    mailbox: "agent@contoso.com", to: ["bob@contoso.com"], subject: "Digest", body: "<p>x</p>", origin: "agentfleet; kind=digest",
  });
  const body = JSON.parse(calls[0].body!);
  assert.deepEqual(body.message.internetMessageHeaders, [{ name: "X-AgentFleet-Origin", value: "agentfleet; kind=digest" }]);
});

test("sendEmail without origin sends no internetMessageHeaders", async () => {
  on("POST", /\/sendMail$/, { status: 202 });
  await provider({ originHeader: "X-AgentFleet-Origin" }).sendEmail({
    mailbox: "agent@contoso.com", to: ["bob@contoso.com"], subject: "x", body: "x",
  });
  assert.equal(JSON.parse(calls[0].body!).message.internetMessageHeaders, undefined);
});

test("origin without originHeader configured throws before any request", async () => {
  await assert.rejects(
    provider().sendEmail({ mailbox: "a@contoso.com", to: ["b@contoso.com"], subject: "x", body: "x", origin: "agentfleet" }),
    /originHeader/,
  );
  assert.equal(calls.length, 0);
});

test("reply with origin goes through createReply with the header at creation, never PATCHed", async () => {
  on("POST", /\/createReply$/, { json: { id: "draft1", body: { contentType: "html", content: "<html><body><p>q</p></body></html>" } } });
  on("GET", /\/attachments$/, { json: { value: [] } });
  on("PATCH", /\/messages\/draft1$/, { json: {} });
  on("POST", /\/draft1\/send$/, { status: 202 });
  await provider({ originHeader: "X-AgentFleet-Origin" }).replyToMessage({
    mailbox: "agent@contoso.com", messageId: "orig1", body: "<p>ok</p>", origin: "agentfleet",
  });
  const create = calls.find((c) => c.url.pathname.endsWith("/createReply"))!;
  assert.deepEqual(JSON.parse(create.body!).message.internetMessageHeaders, [{ name: "X-AgentFleet-Origin", value: "agentfleet" }]);
  const patch = calls.find((c) => c.method === "PATCH")!;
  assert.equal(JSON.parse(patch.body!).internetMessageHeaders, undefined);
  assert.ok(calls.some((c) => c.url.pathname.endsWith("/draft1/send")));
  assert.ok(!calls.some((c) => c.url.pathname.endsWith("/reply")), "one-shot /reply not used");
});

test("reply without origin or attachments keeps the one-shot /reply", async () => {
  on("POST", /\/reply$/, { status: 202 });
  await provider({ originHeader: "X-AgentFleet-Origin" }).replyToMessage({ mailbox: "a@contoso.com", messageId: "o", body: "x" });
  assert.equal(calls.length, 1);
  assert.match(calls[0].url.pathname, /\/messages\/o\/reply$/);
});

test("forward with origin goes through createForward with the header at creation", async () => {
  on("POST", /\/createForward$/, { json: { id: "draft2", body: { contentType: "html", content: "<p>q</p>" } } });
  on("PATCH", /\/messages\/draft2$/, { json: {} });
  on("POST", /\/draft2\/send$/, { status: 202 });
  await provider({ originHeader: "X-AgentFleet-Origin" }).forwardMessage({
    mailbox: "a@contoso.com", messageId: "o", to: ["c@contoso.com"], origin: "agentfleet",
  });
  const create = calls.find((c) => c.url.pathname.endsWith("/createForward"))!;
  assert.deepEqual(JSON.parse(create.body!).message.internetMessageHeaders, [{ name: "X-AgentFleet-Origin", value: "agentfleet" }]);
});

// ── MIME conversation message ──

test("sendConversationMessage: MIME sendMail with our Message-ID, thread, origin and Auto-Submitted", async () => {
  on("POST", /\/sendMail$/, { status: 202 });
  const { messageId } = await provider({ originHeader: "X-AgentFleet-Origin" }).sendConversationMessage({
    mailbox: "agent@contoso.com",
    to: ["bob@contoso.com"],
    cc: ["carol@contoso.com"],
    subject: "Re: Pedido 1234",
    html: '<p>Listo</p><img src="cid:sig">',
    text: "Listo",
    inReplyTo: "<p1@contoso.com>",
    references: "<p0@contoso.com> <p1@contoso.com>",
    autoSubmitted: "auto-replied",
    origin: "agentfleet; kind=reply",
    attachments: [{ filename: "sig.png", contentType: "image/png", content: Buffer.from("png"), contentId: "sig" }],
  });
  assert.match(messageId, /^<[0-9a-f-]{36}@contoso\.com>$/);
  const c = calls[0];
  assert.equal(c.headers["Content-Type"], "text/plain");
  const mime = Buffer.from(c.body!, "base64").toString();
  assert.match(mime, new RegExp(`\r\nMessage-ID: ${messageId}\r\n`));
  assert.match(mime, /\r\nIn-Reply-To: <p1@contoso\.com>\r\nReferences: <p0@contoso\.com> <p1@contoso\.com>\r\n/);
  assert.match(mime, /\r\nX-AgentFleet-Origin: agentfleet; kind=reply\r\n/);
  assert.match(mime, /\r\nAuto-Submitted: auto-replied\r\nX-Auto-Response-Suppress: All\r\n/);
  assert.match(mime, /\r\nCc: carol@contoso\.com\r\n/);
  assert.match(mime, /Content-ID: <sig>/);
  assert.doesNotMatch(mime, /\r\nReply-To:/i);
});

test("sendConversationMessage without autoSubmitted/thread: none of those headers", async () => {
  on("POST", /\/sendMail$/, { status: 202 });
  await provider().sendConversationMessage({ mailbox: "agent@contoso.com", to: ["bob@contoso.com"], subject: "Hola", html: "<p>x</p>" });
  const mime = Buffer.from(calls[0].body!, "base64").toString();
  assert.doesNotMatch(mime, /Auto-Submitted|X-Auto-Response-Suppress|In-Reply-To|References/);
});

// ── Find / move / folders ──

test("findByInternetMessageId: OData-quoted filter, scoped to the folder", async () => {
  on("GET", /\/mailFolders\/sentitems\/messages$/, { json: { value: [fx.reply] } });
  const r = await provider().findByInternetMessageId("agent@contoso.com", "<o'brien@contoso.com>", "sent");
  assert.equal(r.length, 1);
  assert.equal(calls[0].url.searchParams.get("$filter"), "internetMessageId eq '<o''brien@contoso.com>'");
});

test("findByInternetMessageId without folder searches the whole mailbox; [] when not there yet", async () => {
  on("GET", /\/users\/agent@contoso\.com\/messages$/, { json: { value: [] } });
  assert.deepEqual(await provider().findByInternetMessageId("agent@contoso.com", "<x@contoso.com>"), []);
});

test("moveMessage maps neutral names and returns the new id", async () => {
  on("POST", /\/messages\/m1\/move$/, { json: { id: "m1-moved" } });
  assert.deepEqual(await provider().moveMessage("a@contoso.com", "m1", "trash"), { id: "m1-moved" });
  assert.deepEqual(JSON.parse(calls[0].body!), { destinationId: "deleteditems" });
});

const folder = { id: "F1", displayName: "AgentFleet", totalItemCount: 3, unreadItemCount: 1, childFolderCount: 0 };

test("getOrCreateFolder returns an existing (hidden) folder without creating", async () => {
  on("GET", /\/mailFolders$/, { json: { value: [folder] } });
  const f = await provider().getOrCreateFolder("a@contoso.com", "AgentFleet", { hidden: true });
  assert.deepEqual(f, { id: "F1", name: "AgentFleet", totalCount: 3, unreadCount: 1, childCount: 0 });
  assert.equal(calls[0].url.searchParams.get("includeHiddenFolders"), "true");
  assert.equal(calls[0].url.searchParams.get("$filter"), "displayName eq 'AgentFleet'");
  assert.equal(calls.length, 1);
});

test("getOrCreateFolder creates it hidden under the parent when missing", async () => {
  on("GET", /\/childFolders$/, { json: { value: [] } });
  on("POST", /\/mailFolders\/inbox\/childFolders$/, { json: folder });
  await provider().getOrCreateFolder("a@contoso.com", "AgentFleet", { parentId: "inbox", hidden: true });
  assert.deepEqual(JSON.parse(calls[1].body!), { displayName: "AgentFleet", isHidden: true });
});

test("getOrCreateFolder: a 409 create race re-reads the folder", async () => {
  let lookups = 0;
  responses.push((c) => (c.method === "GET" ? { json: { value: lookups++ === 0 ? [] : [folder] } } : undefined));
  on("POST", /\/mailFolders$/, { status: 409, json: { error: { code: "ErrorFolderExists", message: "exists" } } });
  const f = await provider().getOrCreateFolder("a@contoso.com", "AgentFleet");
  assert.equal(f.id, "F1");
});

test("getOrCreateFolder: other errors propagate", async () => {
  on("GET", /\/mailFolders$/, { json: { value: [] } });
  on("POST", /\/mailFolders$/, { status: 403, json: { error: { code: "ErrorAccessDenied", message: "denied" } } });
  await assert.rejects(provider().getOrCreateFolder("a@contoso.com", "AgentFleet"), /Graph API 403/);
});

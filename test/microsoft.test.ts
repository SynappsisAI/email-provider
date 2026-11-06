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

// ── Drafts ──

const sent = () => calls.filter((c) => /\/(send|sendMail)$/.test(c.url.pathname));

test("createDraft: sendEmail's message POSTed to /messages, never sent; returns the draft's ids", async () => {
  on("POST", /\/users\/ana@contoso\.com\/messages$/, {
    json: { id: "D1", internetMessageId: "<d1@contoso.com>", conversationId: "C1", isDraft: true },
  });
  const ref = await provider({ originHeader: "X-AgentFleet-Origin" }).createDraft({
    mailbox: "ana@contoso.com",
    to: ["bob@contoso.com"],
    cc: ["carol@contoso.com"],
    bcc: ["dave@contoso.com"],
    subject: "Propuesta",
    body: '<p>Hola</p><img src="cid:logo">',
    attachments: [{ filename: "logo.png", contentType: "image/png", content: Buffer.from("png"), contentId: "logo" }],
  });
  assert.deepEqual(ref, { id: "D1", draftId: "D1", internetMessageId: "<d1@contoso.com>", conversationId: "C1" });
  assert.equal(calls.length, 1);
  assert.deepEqual(sent(), []);
  const body = JSON.parse(calls[0].body!);
  assert.deepEqual(body.toRecipients, [{ emailAddress: { address: "bob@contoso.com" } }]);
  assert.deepEqual(body.ccRecipients, [{ emailAddress: { address: "carol@contoso.com" } }]);
  assert.deepEqual(body.bccRecipients, [{ emailAddress: { address: "dave@contoso.com" } }]);
  assert.deepEqual(body.body, { contentType: "HTML", content: '<p>Hola</p><img src="cid:logo">' });
  assert.equal(body.attachments[0].contentBytes, Buffer.from("png").toString("base64"));
  assert.equal(body.attachments[0].contentId, "logo");
  assert.equal(body.internetMessageHeaders, undefined, "no origin header unless the caller passes one");
});

test("createDraft: reads internetMessageId back when the create response lacks it", async () => {
  on("POST", /\/messages$/, { json: { id: "D1" } });
  on("GET", /\/messages\/D1$/, { json: { internetMessageId: "<d1@contoso.com>" } });
  const ref = await provider().createDraft({ mailbox: "ana@contoso.com", to: ["b@contoso.com"], subject: "s", body: "b" });
  assert.equal(ref.internetMessageId, "<d1@contoso.com>");
  assert.equal(calls[1].url.searchParams.get("$select"), "internetMessageId");
});

const replyDraft = { id: "D2", internetMessageId: "<d2@contoso.com>", conversationId: "C9", body: { contentType: "html", content: "<html><body><p>quoted</p></body></html>" } };

test("createReplyDraft: Graph's createReply (Reply-To/From rules), body above the quote, never sent", async () => {
  on("POST", /\/messages\/orig1\/createReply$/, { json: replyDraft });
  on("PATCH", /\/messages\/D2$/, { json: {} });
  on("GET", /\/attachments$/, { json: { value: [] } });
  const ref = await provider().createReplyDraft({ mailbox: "ana@contoso.com", messageId: "orig1", body: "<p>ok</p>" });
  assert.deepEqual(ref, { id: "D2", draftId: "D2", internetMessageId: "<d2@contoso.com>", conversationId: "C9" });
  assert.deepEqual(sent(), []);
  assert.ok(!calls.some((c) => /\/reply$/.test(c.url.pathname)));
  const create = calls.find((c) => c.url.pathname.endsWith("/createReply"))!;
  assert.deepEqual(JSON.parse(create.body!), {}, "no origin header, recipients left to Graph");
  const patch = JSON.parse(calls.find((c) => c.method === "PATCH")!.body!);
  assert.equal(patch.body.content, "<html><body><p>ok</p><p>quoted</p></body></html>");
  assert.equal(patch.toRecipients, undefined);
});

test("createReplyDraft replyAll: createReplyAll", async () => {
  on("POST", /\/createReplyAll$/, { json: replyDraft });
  on("PATCH", /\/messages\/D2$/, { json: {} });
  on("GET", /\/attachments$/, { json: { value: [] } });
  await provider().createReplyDraft({ mailbox: "ana@contoso.com", messageId: "orig1", body: "x", replyAll: true });
  assert.ok(calls.some((c) => c.url.pathname.endsWith("/orig1/createReplyAll")));
  assert.deepEqual(sent(), []);
});

test("createForwardDraft: createForward with the given recipients and comment, never sent", async () => {
  on("POST", /\/createForward$/, { json: { ...replyDraft, id: "D3" } });
  on("PATCH", /\/messages\/D3$/, { json: {} });
  on("POST", /\/messages\/D3\/attachments$/, { json: {} });
  const ref = await provider().createForwardDraft({
    mailbox: "ana@contoso.com",
    messageId: "orig1",
    to: ["eve@contoso.com"],
    comment: "<p>FYI</p>",
    attachments: [{ filename: "n.txt", contentType: "text/plain", content: Buffer.from("n") }],
  });
  assert.equal(ref.draftId, "D3");
  assert.deepEqual(sent(), []);
  const patch = JSON.parse(calls.find((c) => c.method === "PATCH")!.body!);
  assert.deepEqual(patch.toRecipients, [{ emailAddress: { address: "eve@contoso.com" } }]);
  assert.match(patch.body.content, /<body><p>FYI<\/p><p>quoted<\/p>/);
  assert.ok(calls.some((c) => c.method === "POST" && c.url.pathname.endsWith("/D3/attachments")));
});

test("createForwardDraft: a failure after creation deletes the draft and never sends", async () => {
  on("POST", /\/createForward$/, { json: { ...replyDraft, id: "D4" } });
  on("PATCH", /\/messages\/D4$/, { json: {} });
  on("POST", /\/D4\/attachments$/, { status: 413, json: { error: { code: "ErrorRequestTooLarge", message: "too large" } } });
  on("DELETE", /\/messages\/D4$/, { status: 204 });
  await assert.rejects(
    provider().createForwardDraft({
      mailbox: "ana@contoso.com",
      messageId: "orig1",
      to: ["eve@contoso.com"],
      attachments: [{ filename: "big.bin", contentType: "application/octet-stream", content: Buffer.from("x") }],
    }),
    /Graph API 413/,
  );
  assert.deepEqual(sent(), []);
  assert.ok(calls.some((c) => c.method === "DELETE" && c.url.pathname.endsWith("/messages/D4")));
});

test("deleteDraft: deletes a draft after the isDraft check", async () => {
  on("GET", /\/messages\/D1$/, { json: { id: "D1", isDraft: true } });
  on("DELETE", /\/messages\/D1$/, { status: 204 });
  await provider().deleteDraft("ana@contoso.com", "D1");
  assert.equal(calls[0].url.searchParams.get("$select"), "id,isDraft");
  assert.deepEqual(calls.map((c) => c.method), ["GET", "DELETE"]);
});

test("deleteDraft: refuses a message that isn't a draft (status 404, like Gmail), deleting nothing", async () => {
  on("GET", /\/messages\/M1$/, { json: { id: "M1", isDraft: false } });
  await assert.rejects(provider().deleteDraft("ana@contoso.com", "M1"), (e: any) => /not a draft/.test(e.message) && e.status === 404);
  assert.ok(!calls.some((c) => c.method === "DELETE"));
});

test("deleteDraft: a lookup failure deletes nothing", async () => {
  on("GET", /\/messages\/X$/, { status: 404, json: { error: { code: "ErrorItemNotFound", message: "not found" } } });
  await assert.rejects(provider().deleteDraft("ana@contoso.com", "X"), /no draft X/);
  assert.ok(!calls.some((c) => c.method === "DELETE"));
});

const attReply = {
  mailbox: "ana@contoso.com",
  messageId: "orig1",
  body: "x",
  attachments: [{ filename: "a.txt", contentType: "text/plain", content: Buffer.from("a") }],
};
const draftFlow = (id: string) => {
  on("POST", /\/createReply$/, { json: { ...replyDraft, id } });
  on("PATCH", new RegExp(`/messages/${id}$`), { json: {} });
  on("GET", /\/attachments$/, { json: { value: [] } });
};

test("createReplyDraft: internetMessageId taken from the PATCH response, no extra GET", async () => {
  on("POST", /\/createReply$/, { json: { id: "D6", body: { contentType: "html", content: "<p>q</p>" } } });
  on("PATCH", /\/messages\/D6$/, { json: { id: "D6", internetMessageId: "<d6@contoso.com>" } });
  on("GET", /\/attachments$/, { json: { value: [] } });
  const ref = await provider().createReplyDraft({ mailbox: "ana@contoso.com", messageId: "orig1", body: "x" });
  assert.equal(ref.internetMessageId, "<d6@contoso.com>");
  assert.ok(!calls.some((c) => c.method === "GET" && c.url.pathname.endsWith("/messages/D6")));
});

test("deleteDraft: a draft that no longer exists gets the same wording as Gmail, status kept", async () => {
  on("GET", /\/messages\/GONE$/, { status: 404, json: { error: { code: "ErrorItemNotFound", message: "not found" } } });
  await assert.rejects(provider().deleteDraft("ana@contoso.com", "GONE"), (err: any) => {
    assert.match(err.message, /no draft GONE in ana@contoso\.com — already sent or discarded/);
    assert.equal(err.status, 404);
    return true;
  });
});

test("a token failure on an ordinary call still surfaces MSAL's own error", async () => {
  const p = provider();
  class MsalLikeError extends Error { errorCode = "invalid_client"; }
  (p as any).getToken = async () => { throw new MsalLikeError("bad secret"); };
  await assert.rejects(p.listFolders("ana@contoso.com"), (err: any) => err instanceof MsalLikeError);
});

test("send via draft: a failed /send deletes the draft (v0.7 behavior; see #12) and re-throws", async () => {
  draftFlow("D5");
  on("POST", /\/D5\/send$/, { status: 503, json: { error: { code: "ServiceUnavailable", message: "busy" } } });
  on("DELETE", /\/messages\/D5$/, { status: 204 });
  await assert.rejects(provider().replyToMessage(attReply), /^Error: Graph API 503: busy$/);
  assert.ok(calls.some((c) => c.method === "DELETE" && c.url.pathname.endsWith("/messages/D5")));
});

test("deleteDraft: an empty id is refused before any request", async () => {
  await assert.rejects(provider().deleteDraft("ana@contoso.com", " "), /draftId is empty/);
  assert.equal(calls.length, 0);
});

// ── Path safety ──

test("crafted ids are encoded into ONE path segment: requests never leave the mailbox", async () => {
  on("GET", /./, { status: 404, json: { error: { code: "ErrorInvalidIdMalformed", message: "bad id" } } });
  on("PATCH", /./, { status: 404, json: { error: { code: "ErrorInvalidIdMalformed", message: "bad id" } } });
  const p = provider();
  const evil = [
    "../../other@contoso.com/messages/AAMkX=",
    "%2e%2e/%2e%2e/other@contoso.com/messages/AAMkX=",
    "..\\..\\other@contoso.com\\messages\\AAMkX=",
    ".\n./.\t./other@contoso.com/messages/AAMkX=",
    ".. ",
    "AAMkX=?$select=body",
    "AAMkX=#frag",
  ];
  for (const id of evil) {
    await assert.rejects(p.markAsRead("ana@contoso.com", id, true));
    await assert.rejects(p.deleteDraft("ana@contoso.com", id));
    await assert.rejects(p.getAttachment("ana@contoso.com", "AAMkX=", id));
  }
  assert.ok(calls.length > 0);
  for (const c of calls) {
    const segs = c.url.pathname.split("/").slice(1); // ["v1.0","users","ana@contoso.com","messages",<id>,...]
    assert.deepEqual(segs.slice(0, 4), ["v1.0", "users", "ana@contoso.com", "messages"], c.url.pathname);
    assert.ok(!segs.includes("..") && !segs.includes("."), c.url.pathname);
    assert.equal(c.url.search.includes("select=body"), false);
  }
  assert.ok(!calls.some((c) => c.method === "DELETE"));
});

test("'.' / '..' / empty ids are refused before any request", async () => {
  const p = provider();
  for (const id of [".", "..", ""]) await assert.rejects(p.readMessage("ana@contoso.com", id), /Invalid mailbox or id/);
  await assert.rejects(p.listFolders(".."), /Invalid mailbox or id/);
  assert.equal(calls.length, 0);
});

test("ordinary paths are unchanged: '@', '=', '-', '_' and quotes stay literal", async () => {
  on("PATCH", /./, { json: {} });
  await provider().markAsRead("o'brien@contoso.com", "AAMkAGI2-T_VSAAA=", true);
  assert.equal(calls[0].url.pathname, "/v1.0/users/o'brien@contoso.com/messages/AAMkAGI2-T_VSAAA=");
});

test("createDraft: a create response without an id is an error, not a DraftRef of undefineds", async () => {
  on("POST", /\/messages$/, { json: {} });
  await assert.rejects(provider().createDraft({ mailbox: "ana@contoso.com", to: ["b@contoso.com"], subject: "s", body: "b" }), /no id/);
});

test("deleteDraft: a malformed id gets the cross-provider 'no draft' wording, status kept", async () => {
  on("GET", /./, { status: 400, json: { error: { code: "ErrorInvalidIdMalformed", message: "Id is malformed." } } });
  await assert.rejects(provider().deleteDraft("ana@contoso.com", "r-1234"), (e: any) => /no draft r-1234 .* not a draft id/.test(e.message) && e.status === 400);
  assert.ok(!calls.some((c) => c.method === "DELETE"));
});

test("createReplyDraft: a create response without an id says so (not 'invalid id' blaming the caller)", async () => {
  on("POST", /\/createReply$/, { json: {} });
  await assert.rejects(provider().createReplyDraft({ mailbox: "ana@contoso.com", messageId: "orig1", body: "x" }), /Graph returned no id for the createReply draft/);
});

test("deleteDraft: a DELETE that fails after the check is reported, never swallowed", async () => {
  on("GET", /\/messages\/D1$/, { json: { id: "D1", isDraft: true } });
  on("DELETE", /\/messages\/D1$/, { status: 503, json: { error: { code: "ServiceUnavailable", message: "busy" } } });
  await assert.rejects(provider().deleteDraft("ana@contoso.com", "D1"), (e: any) => e.status === 503);
});

test("deleteDraft: an item-not-found 404 on the DELETE reports 'no draft'", async () => {
  on("GET", /\/messages\/D1$/, { json: { id: "D1", isDraft: true } });
  on("DELETE", /\/messages\/D1$/, { status: 404, json: { error: { code: "ErrorItemNotFound", message: "not found" } } });
  await assert.rejects(provider().deleteDraft("ana@contoso.com", "D1"), (e: any) => /no draft D1 .* already sent or discarded/.test(e.message) && e.status === 404);
});

test("deleteDraft: a 404 about the mailbox is not reported as 'already sent or discarded'", async () => {
  on("GET", /./, { status: 404, json: { error: { code: "ErrorInvalidUser", message: "no such user" } } });
  await assert.rejects(provider().deleteDraft("typo@contoso.com", "D1"), /^Error: Graph API 404: no such user$/);
});

// ── sendDraft ──

test("sendDraft: checks isDraft, then POSTs /send with no body, and touches nothing else", async () => {
  on("GET", /\/messages\/D1$/, { json: { id: "D1", isDraft: true } });
  on("POST", /\/messages\/D1\/send$/, { status: 202 });
  await provider().sendDraft("ana@contoso.com", "D1");
  assert.deepEqual(calls.map((c) => [c.method, c.url.pathname]), [
    ["GET", "/v1.0/users/ana@contoso.com/messages/D1"],
    ["POST", "/v1.0/users/ana@contoso.com/messages/D1/send"],
  ]);
  assert.equal(calls[1].body, undefined);
});

test("sendDraft: refuses a message that isn't a draft, sending nothing", async () => {
  on("GET", /\/messages\/M1$/, { json: { id: "M1", isDraft: false } });
  await assert.rejects(provider().sendDraft("ana@contoso.com", "M1"), (e: any) => /^sendDraft: message M1 is not a draft/.test(e.message) && e.status === 404);
  assert.equal(calls.length, 1);
});

test("sendDraft: a failed send keeps the draft and the error's status (it may have gone out)", async () => {
  on("GET", /\/messages\/D1$/, { json: { id: "D1", isDraft: true } });
  on("POST", /\/D1\/send$/, { status: 503, json: { error: { code: "ServiceUnavailable", message: "busy" } } });
  await assert.rejects(provider().sendDraft("ana@contoso.com", "D1"), (e: any) => e.status === 503);
  assert.ok(!calls.some((c) => c.method === "DELETE"));
});

test("sendDraft: an id that doesn't exist or is malformed is refused at the check, sending nothing", async () => {
  on("GET", /\/messages\/GONE$/, { status: 404, json: { error: { code: "ErrorItemNotFound", message: "not found" } } });
  on("GET", /\/messages\/r-1$/, { status: 400, json: { error: { code: "ErrorInvalidIdMalformed", message: "Id is malformed." } } });
  await assert.rejects(provider().sendDraft("ana@contoso.com", "GONE"), /^Error: sendDraft: no draft GONE in ana@contoso\.com — already sent or discarded$/);
  await assert.rejects(provider().sendDraft("ana@contoso.com", "r-1"), (e: any) => /no draft r-1 .* not a draft id/.test(e.message) && e.status === 400);
  assert.ok(!calls.some((c) => c.method === "POST"));
});

test("sendDraft: a 404 from /send stays Graph's error, not 'already sent' (the check just saw the draft)", async () => {
  on("GET", /\/messages\/D1$/, { json: { id: "D1", isDraft: true } });
  on("POST", /\/D1\/send$/, { status: 404, json: { error: { code: "ErrorItemNotFound", message: "not found" } } });
  await assert.rejects(provider().sendDraft("ana@contoso.com", "D1"), /^Error: Graph API 404: not found$/);
});

test("sendDraft: sending again after success is refused under immutableIds (same id, no longer a draft)", async () => {
  let sent = false;
  responses.push((c) => (c.method === "GET" ? { json: { id: "D1", isDraft: !sent } } : undefined));
  responses.push((c) => {
    if (c.method !== "POST") return undefined;
    sent = true;
    return { status: 202 };
  });
  const p = provider({ immutableIds: true });
  await p.sendDraft("ana@contoso.com", "D1");
  await assert.rejects(p.sendDraft("ana@contoso.com", "D1"), /not a draft/);
  assert.equal(calls.filter((c) => c.method === "POST").length, 1);
});

test("sendDraft: an empty id is refused before any request; a crafted id stays in the mailbox", async () => {
  await assert.rejects(provider().sendDraft("ana@contoso.com", " "), /sendDraft: draftId is empty/);
  assert.equal(calls.length, 0);
  on("GET", /./, { status: 404, json: { error: { code: "ErrorItemNotFound", message: "not found" } } });
  await assert.rejects(provider().sendDraft("ana@contoso.com", "../../other@contoso.com/messages/X"));
  assert.ok(calls[0].url.pathname.startsWith("/v1.0/users/ana@contoso.com/messages/"));
  assert.ok(!calls.some((c) => c.method === "POST"));
});

import { test } from "node:test";
import assert from "node:assert/strict";
import { GoogleEmailProvider } from "../src/providers/google.js";

type Call = { api: string; args: any };

/** A provider whose Gmail client for every mailbox is a recording fake. */
function provider(handlers: Record<string, (args: any) => any> = {}, originHeader?: string) {
  const calls: Call[] = [];
  const api = (name: string) => async (args: any) => {
    calls.push({ api: name, args });
    return { data: (handlers[name] ?? (() => ({})))(args) };
  };
  const fake = {
    users: {
      messages: {
        list: api("messages.list"),
        get: api("messages.get"),
        send: api("messages.send"),
        modify: api("messages.modify"),
        trash: api("messages.trash"),
        attachments: { get: api("messages.attachments.get") },
      },
      drafts: { create: api("drafts.create"), delete: api("drafts.delete"), send: api("drafts.send") },
      labels: { list: api("labels.list"), get: api("labels.get"), create: api("labels.create") },
    },
  };
  const p = new GoogleEmailProvider({ type: "google", serviceAccountKey: {}, ...(originHeader ? { originHeader } : {}) });
  (p as any).getGmail = () => fake;
  return { p, calls };
}

const H = (...pairs: [string, string][]) => pairs.map(([name, value]) => ({ name, value }));

const toFull = (m: unknown, originHeader?: string) => (GoogleEmailProvider as any).toFull(m, originHeader);

test("toFull: threadId as conversationId, Message-ID, origin, kind; Microsoft-only fields undefined", () => {
  const m = toFull(
    {
      id: "g1",
      threadId: "t1",
      internalDate: "1760367600000",
      payload: {
        mimeType: "text/html",
        headers: H(["From", "Bob <bob@example.com>"], ["Message-ID", "<m1@example.com>"], ["X-AgentFleet-Origin", "agentfleet"]),
        body: { data: Buffer.from("<p>x</p>").toString("base64url") },
      },
    },
    "X-AgentFleet-Origin",
  );
  assert.equal(m.conversationId, "t1");
  assert.equal(m.internetMessageId, "<m1@example.com>");
  assert.equal(m.origin, "agentfleet");
  assert.equal(m.kind, "message");
  assert.equal(m.parentFolderId, undefined);
  assert.equal(m.uniqueBodyHtml, undefined);
  assert.equal(m.internalSender, undefined);
});

test("toFull: a text/calendar part makes it a meeting", () => {
  const m = toFull({
    id: "g2",
    internalDate: "0",
    payload: {
      mimeType: "multipart/mixed",
      headers: H(["From", "cal@example.com"]),
      parts: [{ mimeType: "multipart/alternative", parts: [{ mimeType: "text/calendar", body: {} }] }],
    },
  });
  assert.equal(m.kind, "meeting");
});

test("toFull: an attached .ics is not a meeting", () => {
  const ics = (extra: object) =>
    toFull({
      id: "g3",
      internalDate: "0",
      payload: {
        mimeType: "multipart/mixed",
        headers: H(["From", "bob@example.com"]),
        parts: [{ mimeType: "text/html", body: { data: "" } }, { mimeType: "text/calendar", body: { attachmentId: "a1" }, ...extra }],
      },
    }).kind;
  assert.equal(ics({ filename: "evento.ics" }), "message");
  assert.equal(ics({ headers: H(["Content-Disposition", "attachment"]) }), "message");
});

test("listMessages: summaries ask for Cc and Message-ID and carry them", async () => {
  const { p, calls } = provider({
    "messages.list": () => ({ messages: [{ id: "g1" }] }),
    "messages.get": () => ({
      id: "g1",
      threadId: "t1",
      internalDate: "0",
      payload: { headers: H(["From", "a@example.com"], ["Cc", "c@example.com"], ["Message-ID", "<m1@example.com>"]) },
    }),
  });
  const r = await p.listMessages({ mailbox: "agent@example.com" });
  assert.deepEqual(r.messages[0].cc, [{ address: "c@example.com" }]);
  assert.equal(r.messages[0].conversationId, "t1");
  assert.equal(r.messages[0].internetMessageId, "<m1@example.com>");
  assert.deepEqual(calls.find((c) => c.api === "messages.get")!.args.metadataHeaders, ["From", "To", "Cc", "Subject", "Message-ID"]);
});

test("findByInternetMessageId: rfc822msgid without brackets, scoped to the folder label", async () => {
  const { p, calls } = provider({ "messages.list": () => ({}) });
  assert.deepEqual(await p.findByInternetMessageId("a@example.com", "<m1@example.com>", "sent"), []);
  assert.equal(calls[0].args.q, "rfc822msgid:m1@example.com");
  assert.deepEqual(calls[0].args.labelIds, ["SENT"]);
});

test("findByInternetMessageId strips characters that could form search operators", async () => {
  const { p, calls } = provider({ "messages.list": () => ({}) });
  await p.findByInternetMessageId("a@example.com", "<x)OR(subject:pay roll@example.com>");
  assert.equal(calls[0].args.q, "rfc822msgid:xORsubjectpayroll@example.com");
  assert.deepEqual(await p.findByInternetMessageId("a@example.com", "<()>"), []);
});

test("sendConversationMessage: a search hit whose Message-ID differs is not used as the thread", async () => {
  const { p, calls } = provider({
    "messages.list": () => ({ messages: [{ id: "other" }] }),
    "messages.get": (a) =>
      a.id === "other"
        ? { id: "other", threadId: "unrelated", internalDate: "0", payload: { headers: H(["Message-ID", "<different@example.com>"]) } }
        : { id: "sent1", payload: { headers: [] } },
    "messages.send": () => ({ id: "sent1" }),
  });
  await p.sendConversationMessage({ mailbox: "a@example.com", to: ["b@example.com"], subject: "Re: x", html: "x", inReplyTo: "<p1@example.com>" });
  assert.equal(calls.find((c) => c.api === "messages.send")!.args.requestBody.threadId, undefined);
});

test("sendConversationMessage: threads via the parent's threadId and returns the sent Message-ID", async () => {
  const { p, calls } = provider(
    {
      "messages.list": () => ({ messages: [{ id: "parent" }] }),
      "messages.get": (a) =>
        a.id === "parent"
          ? { id: "parent", threadId: "t9", internalDate: "0", payload: { headers: H(["Message-ID", "<p1@example.com>"]) } }
          : { id: "sent1", payload: { headers: H(["Message-ID", "<gmail-kept@example.com>"]) } },
      "messages.send": () => ({ id: "sent1" }),
    },
    "X-AgentFleet-Origin",
  );
  const r = await p.sendConversationMessage({
    mailbox: "agent@example.com",
    to: ["bob@example.com"],
    subject: "Re: x",
    html: "<p>x</p>",
    inReplyTo: "<p1@example.com>",
    autoSubmitted: "auto-replied",
    origin: "agentfleet",
  });
  const send = calls.find((c) => c.api === "messages.send")!;
  assert.equal(send.args.requestBody.threadId, "t9");
  const mime = Buffer.from(send.args.requestBody.raw, "base64url").toString();
  assert.match(mime, /\r\nIn-Reply-To: <p1@example\.com>\r\n/);
  assert.match(mime, /\r\nX-AgentFleet-Origin: agentfleet\r\n/);
  assert.match(mime, /\r\nAuto-Submitted: auto-replied\r\n/);
  assert.match(mime, /\r\nMessage-ID: <[0-9a-f-]{36}@example\.com>\r\n/);
  assert.equal(r.messageId, "<gmail-kept@example.com>");
});

test("sendEmail/replyToMessage: origin stamped as the configured header", async () => {
  const { p, calls } = provider(
    { "messages.get": () => ({ threadId: "t1", payload: { headers: H(["From", "bob@example.com"], ["Message-ID", "<o@example.com>"]) } }) },
    "X-AgentFleet-Origin",
  );
  await p.sendEmail({ mailbox: "a@example.com", to: ["b@example.com"], subject: "s", body: "b", origin: "agentfleet; kind=digest" });
  await p.replyToMessage({ mailbox: "a@example.com", messageId: "o", body: "b", origin: "agentfleet; kind=reply" });
  const raws = calls.filter((c) => c.api === "messages.send").map((c) => Buffer.from(c.args.requestBody.raw, "base64url").toString());
  assert.match(raws[0], /\r\nX-AgentFleet-Origin: agentfleet; kind=digest\r\n/);
  assert.match(raws[1], /\r\nX-AgentFleet-Origin: agentfleet; kind=reply\r\n/);
});

test("moveMessage: adds the destination label, removes the other location labels; trash uses trash", async () => {
  const { p, calls } = provider();
  assert.deepEqual(await p.moveMessage("a@example.com", "m1", "Label_7"), { id: "m1" });
  assert.deepEqual(calls[0].args.requestBody, { addLabelIds: ["Label_7"], removeLabelIds: ["INBOX", "SPAM", "TRASH"] });
  await p.moveMessage("a@example.com", "m1", "inbox");
  assert.deepEqual(calls[1].args.requestBody, { addLabelIds: ["INBOX"], removeLabelIds: ["SPAM", "TRASH"] });
  await p.moveMessage("a@example.com", "m1", "trash");
  assert.equal(calls[2].api, "messages.trash");
});

test("getOrCreateFolder: existing label (case-insensitive) or a new hidden nested one", async () => {
  const existing = provider({
    "labels.list": () => ({ labels: [{ id: "Label_1", name: "AgentFleet" }] }),
    "labels.get": () => ({ id: "Label_1", name: "AgentFleet", messagesTotal: 2, messagesUnread: 1 }),
  });
  assert.deepEqual(await existing.p.getOrCreateFolder("a@example.com", "agentfleet"), {
    id: "Label_1", name: "AgentFleet", totalCount: 2, unreadCount: 1, childCount: 0,
  });
  assert.ok(!existing.calls.some((c) => c.api === "labels.create"));

  const fresh = provider({
    "labels.get": () => ({ id: "Label_P", name: "Parent" }),
    "labels.list": () => ({ labels: [] }),
    "labels.create": (a) => ({ id: "Label_2", name: a.requestBody.name }),
  });
  const f = await fresh.p.getOrCreateFolder("a@example.com", "Hidden", { parentId: "Label_P", hidden: true });
  assert.equal(f.name, "Parent/Hidden");
  assert.deepEqual(fresh.calls.find((c) => c.api === "labels.create")!.args.requestBody, {
    name: "Parent/Hidden", labelListVisibility: "labelHide", messageListVisibility: "hide",
  });
});

// ── Drafts ──

const SENDS = ["messages.send", "drafts.send"];
const draftMime = (calls: Call[]) =>
  Buffer.from(calls.find((c) => c.api === "drafts.create")!.args.requestBody.message.raw, "base64url").toString();
/** Header block of a MIME string, unfolded into "Name: value" lines. */
const headerOf = (mime: string, name: string) =>
  mime.split("\r\n\r\n")[0].split("\r\n").find((l) => l.toLowerCase().startsWith(`${name.toLowerCase()}:`))?.slice(name.length + 1).trim();

/** Gmail fake for drafts: `original` answers gets of message "o"; the draft is "dm1" / "r-1". */
function draftProvider(original: gmail_message = { threadId: "t1", payload: { headers: [] } }, extra: Record<string, (a: any) => any> = {}) {
  return provider({
    "messages.get": (a) =>
      a.id === "dm1" ? { id: "dm1", payload: { headers: H(["Message-ID", "<kept@example.com>"]) } } : original,
    "drafts.create": () => ({ id: "r-1", message: { id: "dm1", threadId: "t1" } }),
    ...extra,
  });
}
type gmail_message = Record<string, any>;

test("createDraft: sendEmail's MIME saved via drafts.create, never sent; returns the draft's ids", async () => {
  const { p, calls } = draftProvider();
  const ref = await p.createDraft({
    mailbox: "ana@example.com",
    to: ["bob@example.com"],
    cc: ["carol@example.com"],
    bcc: ["dave@example.com"],
    subject: "Propuesta",
    body: '<p>Hola</p><img src="cid:logo">',
    attachments: [
      { filename: "logo.png", contentType: "image/png", content: Buffer.from("png"), contentId: "logo" },
      { filename: "a.pdf", contentType: "application/pdf", content: Buffer.from("pdf") },
    ],
  });
  const mime = draftMime(calls);
  assert.deepEqual(ref, { id: "dm1", draftId: "r-1", internetMessageId: headerOf(mime, "Message-ID"), conversationId: "t1" });
  assert.deepEqual(calls.map((c) => c.api), ["drafts.create"], "one call: no send, no read-back");
  assert.equal(headerOf(mime, "From"), "ana@example.com");
  assert.equal(headerOf(mime, "To"), "bob@example.com");
  assert.equal(headerOf(mime, "Cc"), "carol@example.com");
  assert.equal(headerOf(mime, "Bcc"), "dave@example.com");
  assert.match(headerOf(mime, "Message-ID")!, /^<[0-9a-f-]{36}@example\.com>$/);
  assert.match(mime, /Content-ID: <logo>/);
  assert.match(mime, /filename="a\.pdf"/);
});

test("createDraft: no origin header unless the caller passes one", async () => {
  const { p, calls } = provider({ "drafts.create": () => ({ id: "r-1", message: { id: "dm1" } }) }, "X-AgentFleet-Origin");
  await p.createDraft({ mailbox: "a@example.com", to: ["b@example.com"], subject: "s", body: "b" });
  assert.equal(headerOf(draftMime(calls), "X-AgentFleet-Origin"), undefined);
});


const original = (...headers: [string, string][]) => ({
  threadId: "t1",
  payload: { headers: H(["Subject", "Pedido"], ["Message-ID", "<o1@example.com>"], ["References", "<o0@example.com>"], ...headers) },
});

test("createReplyDraft: goes to Reply-To when present, threaded, never sent", async () => {
  const { p, calls } = draftProvider(original(["From", "Bob <bob@example.com>"], ["Reply-To", "Soporte <help@example.com>"]));
  const ref = await p.createReplyDraft({ mailbox: "ana@example.com", messageId: "o", body: "<p>ok</p>" });
  assert.equal(ref.draftId, "r-1");
  assert.ok(!calls.some((c) => SENDS.includes(c.api)));
  const create = calls.find((c) => c.api === "drafts.create")!;
  assert.equal(create.args.requestBody.message.threadId, "t1");
  const mime = draftMime(calls);
  assert.equal(headerOf(mime, "To"), "help@example.com");
  assert.equal(headerOf(mime, "Cc"), undefined);
  assert.equal(headerOf(mime, "Subject"), "Re: Pedido");
  assert.equal(headerOf(mime, "In-Reply-To"), "<o1@example.com>");
  assert.equal(headerOf(mime, "References"), "<o0@example.com> <o1@example.com>");
});

test("createReplyDraft: goes to From without Reply-To", async () => {
  const { p, calls } = draftProvider(original(["From", "Bob <bob@example.com>"], ["To", "ana@example.com"]));
  await p.createReplyDraft({ mailbox: "ana@example.com", messageId: "o", body: "x" });
  assert.equal(headerOf(draftMime(calls), "To"), "bob@example.com");
});

test("createReplyDraft replyAll: sender + original To, original Cc, minus the mailbox itself", async () => {
  const { p, calls } = draftProvider(
    original(
      ["From", "bob@example.com"],
      ["To", "Ana <ana@example.com>, carol@example.com"],
      ["Cc", "dave@example.com, ANA@example.com"],
    ),
  );
  await p.createReplyDraft({ mailbox: "ana@example.com", messageId: "o", body: "x", replyAll: true });
  const mime = draftMime(calls);
  assert.equal(headerOf(mime, "To"), "bob@example.com, carol@example.com");
  assert.equal(headerOf(mime, "Cc"), "dave@example.com");
});

test("createForwardDraft: given recipients, comment above the original, original attachments carried, never sent", async () => {
  const { p, calls } = draftProvider(
    {
      threadId: "t1",
      payload: {
        mimeType: "multipart/mixed",
        headers: H(["Subject", "Factura"], ["From", "bob@example.com"]),
        parts: [
          { mimeType: "text/html", body: { data: Buffer.from("<p>original</p>").toString("base64url") } },
          { mimeType: "application/pdf", filename: "f.pdf", body: { attachmentId: "att1", size: 3 } },
        ],
      },
    },
    { "messages.attachments.get": () => ({ data: Buffer.from("pdf").toString("base64url"), size: 3 }) },
  );
  await p.createForwardDraft({ mailbox: "ana@example.com", messageId: "o", to: ["eve@example.com"], comment: "<p>FYI</p>" });
  assert.ok(!calls.some((c) => SENDS.includes(c.api)));
  const mime = draftMime(calls);
  assert.equal(headerOf(mime, "To"), "eve@example.com");
  assert.equal(headerOf(mime, "Subject"), "Fwd: Factura");
  assert.match(mime, /filename="f\.pdf"/);
  const html = Buffer.from(mime.match(/text\/html; charset="UTF-8"\r\nContent-Transfer-Encoding: base64\r\n\r\n([A-Za-z0-9+/=\r\n]+)/)![1], "base64").toString();
  assert.ok(html.indexOf("FYI") < html.indexOf("Forwarded message") && html.indexOf("Forwarded message") < html.indexOf("original"));
});

test("draft creation failing never falls back to sending", async () => {
  const { p, calls } = draftProvider(original(["From", "bob@example.com"]), {
    "drafts.create": () => {
      throw Object.assign(new Error("quota"), { code: 429 });
    },
  });
  await assert.rejects(p.createReplyDraft({ mailbox: "ana@example.com", messageId: "o", body: "x" }), /quota/);
  await assert.rejects(p.createDraft({ mailbox: "ana@example.com", to: ["b@example.com"], subject: "s", body: "b" }), /quota/);
  assert.ok(!calls.some((c) => SENDS.includes(c.api)));
});

test("deleteDraft: deletes through the drafts API only", async () => {
  const { p, calls } = provider();
  await p.deleteDraft("ana@example.com", "r-1");
  assert.deepEqual(calls.map((c) => [c.api, c.args.id]), [["drafts.delete", "r-1"]]);
});

test("deleteDraft: an id that isn't a draft (404, or 400 invalidArgument) is refused", async () => {
  for (const err of [
    { status: 404, code: "ERR_BAD_REQUEST" }, // gaxios 7 shape: numeric status, string code
    { code: 404 },
    { message: "Invalid draft id", status: 400, response: { status: 400, data: { error: { errors: [{ reason: "invalidArgument", message: "Invalid id value" }] } } } },
  ]) {
    const { p, calls } = provider({
      "drafts.delete": () => {
        throw Object.assign(new Error("nope"), err);
      },
    });
    await assert.rejects(p.deleteDraft("ana@example.com", "18c0ffee"), (e: any) => /not a draft/.test(e.message) && typeof e.status === "number");
    assert.ok(!calls.some((c) => c.api.startsWith("messages.")));
  }
});

test("deleteDraft: other errors propagate untranslated", async () => {
  const { p } = provider({
    "drafts.delete": () => {
      throw Object.assign(new Error("Backend Error"), { code: 500 });
    },
  });
  await assert.rejects(p.deleteDraft("ana@example.com", "r-1"), /Backend Error/);
});

test("deleteDraft: a 400 that isn't about the id (auth/delegation) propagates as is", async () => {
  const { p } = provider({
    "drafts.delete": () => {
      throw Object.assign(new Error("invalid_grant"), { status: 400, response: { data: { error: "invalid_grant" } } });
    },
  });
  await assert.rejects(p.deleteDraft("ana@example.com", "r-1"), /invalid_grant/);
});

test("createDraft: a create response without a message id deletes the draft it made, then throws", async () => {
  const { p, calls } = provider({ "drafts.create": () => ({ id: "r-9" }) });
  await assert.rejects(p.createDraft({ mailbox: "a@example.com", to: ["b@example.com"], subject: "s", body: "b" }), /no draft or message id/);
  assert.deepEqual(calls.filter((c) => c.api === "drafts.delete").map((c) => c.args.id), ["r-9"]);
});

test("replies/forwards don't stack prefixes on Outlook-style \"RE:\" / \"FW:\" subjects", async () => {
  const { p, calls } = draftProvider({ threadId: "t1", payload: { headers: H(["From", "bob@example.com"], ["Subject", "RE: Cotización"]) } });
  await p.createReplyDraft({ mailbox: "ana@example.com", messageId: "o", body: "x" });
  assert.equal(headerOf(draftMime(calls), "Subject"), "=?UTF-8?B?" + Buffer.from("RE: Cotización").toString("base64") + "?=");
  const fw = draftProvider({ threadId: "t1", payload: { mimeType: "text/html", headers: H(["Subject", "FW: x"]), body: { data: "" } } });
  await fw.p.createForwardDraft({ mailbox: "ana@example.com", messageId: "o", to: ["e@example.com"] });
  assert.equal(headerOf(draftMime(fw.calls), "Subject"), "FW: x");
});

test("replyAll: an address in Reply-To and To/Cc appears once; To wins over Cc", async () => {
  const { p, calls } = draftProvider(
    original(["From", "bob@example.com"], ["Reply-To", "team@example.com"], ["To", "TEAM@example.com, ana@example.com, carol@example.com"], ["Cc", "carol@example.com, dave@example.com"]),
  );
  await p.createReplyDraft({ mailbox: "ana@example.com", messageId: "o", body: "x", replyAll: true });
  const mime = draftMime(calls);
  assert.equal(headerOf(mime, "To"), "team@example.com, carol@example.com");
  assert.equal(headerOf(mime, "Cc"), "dave@example.com");
});

test("a reply with no usable recipient is refused before any draft is created", async () => {
  const { p, calls } = draftProvider(original(["From", ""]));
  await assert.rejects(p.createReplyDraft({ mailbox: "ana@example.com", messageId: "o", body: "x" }), /no recipient/);
  assert.ok(!calls.some((c) => c.api === "drafts.create" || SENDS.includes(c.api)));
});

test("a plain reply to the mailbox's own sent message goes back to the mailbox (as before)", async () => {
  const { p, calls } = draftProvider(original(["From", "Ana <ana@example.com>"], ["To", "bob@example.com"]));
  await p.createReplyDraft({ mailbox: "ana@example.com", messageId: "o", body: "x" });
  assert.equal(headerOf(draftMime(calls), "To"), "ana@example.com");
});

test("reply-all with only Cc left (own message to self) still goes out to the Cc", async () => {
  const { p, calls } = draftProvider(original(["From", "ana@example.com"], ["To", "ana@example.com"], ["Cc", "a@example.com, b@example.com"]));
  await p.createReplyDraft({ mailbox: "ana@example.com", messageId: "o", body: "x", replyAll: true });
  const mime = draftMime(calls);
  assert.equal(headerOf(mime, "To"), "a@example.com, b@example.com", "Cc promoted to To, no empty To header");
  assert.equal(headerOf(mime, "Cc"), undefined);
});

test("deleteDraft: a non-array `errors` on the error doesn't mask the real error", async () => {
  const { p } = provider({
    "drafts.delete": () => {
      throw Object.assign(new Error("Backend Error"), { status: 400, errors: { weird: true } });
    },
  });
  await assert.rejects(p.deleteDraft("ana@example.com", "r-1"), /Backend Error/);
});

test("non-addresses (MAILER-DAEMON, undisclosed-recipients) are not recipients", async () => {
  const bad = draftProvider(original(["From", "MAILER-DAEMON"]));
  await assert.rejects(bad.p.createReplyDraft({ mailbox: "ana@example.com", messageId: "o", body: "x" }), /no recipient/);
  const all = draftProvider(original(["From", "bob@example.com"], ["To", "undisclosed-recipients:;"]));
  await all.p.createReplyDraft({ mailbox: "ana@example.com", messageId: "o", body: "x", replyAll: true });
  assert.equal(headerOf(draftMime(all.calls), "To"), "bob@example.com");
});

test("deleteDraft: another invalidArgument 400 isn't reported as 'no draft'; empty id refused up front", async () => {
  const { p, calls } = provider({
    "drafts.delete": () => {
      throw Object.assign(new Error("Precondition check failed."), { status: 400, errors: [{ reason: "invalidArgument" }] });
    },
  });
  await assert.rejects(p.deleteDraft("ana@example.com", "r-1"), /^Error: Precondition check failed\.$/);
  await assert.rejects(p.deleteDraft("ana@example.com", "  "), /draftId is empty/);
  assert.equal(calls.length, 1);
});

test("a Reply-To with no usable address falls back to From", async () => {
  const { p, calls } = draftProvider(original(["From", "bob@example.com"], ["Reply-To", "list-name"]));
  await p.createReplyDraft({ mailbox: "ana@example.com", messageId: "o", body: "x" });
  assert.equal(headerOf(draftMime(calls), "To"), "bob@example.com");
});

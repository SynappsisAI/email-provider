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
      },
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

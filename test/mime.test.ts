import { test } from "node:test";
import assert from "node:assert/strict";
import { buildMime, buildMimeMessage } from "../src/mime.js";

const logo = { filename: "logo.png", contentType: "image/png", content: Buffer.from([0x89, 0x50, 0x4e, 0x47]), contentId: "logo1" };
const pdf = { filename: "report.pdf", contentType: "application/pdf", content: Buffer.from("%PDF-1.4") };

/** Top-level headers (before the first blank line), unfolded, as a name → values map. */
function headersOf(mime: string): Map<string, string[]> {
  const out = new Map<string, string[]>();
  for (const line of mime.split("\r\n\r\n")[0].split("\r\n")) {
    const i = line.indexOf(":");
    const name = line.slice(0, i).toLowerCase();
    out.set(name, [...(out.get(name) ?? []), line.slice(i + 1).trim()]);
  }
  return out;
}

/** Decoded content of every base64 part with the given Content-Type prefix. */
function partsOf(mime: string, type: string): Buffer[] {
  const out: Buffer[] = [];
  const re = new RegExp(`Content-Type: ${type}[^\\r]*\\r\\n(?:[^\\r]+\\r\\n)*\\r\\n([A-Za-z0-9+/=\\r\\n]+?)\\r\\n--`, "g");
  for (const m of mime.matchAll(re)) out.push(Buffer.from(m[1].replace(/\r\n/g, ""), "base64"));
  return out;
}

test("thread headers, our Message-ID and extra headers", () => {
  const mime = buildMimeMessage({
    from: "agent@contoso.com",
    to: ["bob@contoso.com"],
    cc: ["carol@contoso.com"],
    subject: "Re: Pedido",
    html: "<p>Hola</p>",
    messageId: "<m1@contoso.com>",
    inReplyTo: "<p1@contoso.com>",
    references: "<p0@contoso.com> <p1@contoso.com>",
    headers: [["X-AgentFleet-Origin", "agentfleet; kind=reply"], ["Auto-Submitted", "auto-replied"]],
  });
  const h = headersOf(mime);
  assert.deepEqual(h.get("message-id"), ["<m1@contoso.com>"]);
  assert.deepEqual(h.get("in-reply-to"), ["<p1@contoso.com>"]);
  assert.deepEqual(h.get("references"), ["<p0@contoso.com> <p1@contoso.com>"]);
  assert.deepEqual(h.get("x-agentfleet-origin"), ["agentfleet; kind=reply"]);
  assert.deepEqual(h.get("auto-submitted"), ["auto-replied"]);
  assert.deepEqual(h.get("cc"), ["carol@contoso.com"]);
});

test("no inReplyTo/references → no thread headers (a new thread)", () => {
  const h = headersOf(buildMimeMessage({ from: "a@x.com", to: ["b@x.com"], subject: "Hola", html: "<p>x</p>" }));
  assert.equal(h.has("in-reply-to"), false);
  assert.equal(h.has("references"), false);
});

test("CR/LF in caller values cannot inject headers", () => {
  const mime = buildMimeMessage({
    from: "a@x.com",
    to: ["b@x.com\r\nBcc: evil@y.com"],
    subject: "Hola\r\nBcc: evil@y.com",
    html: "<p>x</p>",
    inReplyTo: "<p@x.com>\r\nX-Injected: 1",
    headers: [["X-Origin", "v\nX-Injected: 2"]],
  });
  const h = headersOf(mime);
  assert.equal(h.has("bcc"), false);
  assert.equal(h.has("x-injected"), false);
});

test("text alternative + inline cid + attachment: mixed > alternative > related", () => {
  const mime = buildMimeMessage({
    from: "a@x.com",
    to: ["b@x.com"],
    subject: "Firma",
    html: '<p>Hola</p><img src="cid:logo1">',
    text: "Hola",
    attachments: [logo, pdf],
  });
  const mixed = mime.indexOf("multipart/mixed");
  const alt = mime.indexOf("multipart/alternative");
  const related = mime.indexOf('multipart/related; type="text/html"');
  assert.ok(mixed >= 0 && mixed < alt && alt < related, "nesting order");
  assert.match(mime, /Content-ID: <logo1>\r\nContent-Disposition: inline; filename="logo.png"/);
  assert.match(mime, /Content-Disposition: attachment; filename="report.pdf"/);
  assert.equal(partsOf(mime, "text/plain")[0].toString(), "Hola");
  assert.equal(partsOf(mime, "text/html")[0].toString(), '<p>Hola</p><img src="cid:logo1">');
  assert.deepEqual(partsOf(mime, "image/png")[0], logo.content);
});

test("non-ASCII subject is RFC 2047 encoded, body is base64 UTF-8", () => {
  const mime = buildMimeMessage({ from: "a@x.com", to: ["b@x.com"], subject: "Revisión", html: "<p>¿Sí?</p>" });
  assert.match(mime, /Subject: =\?UTF-8\?B\?/);
  assert.equal(partsOf(mime + "\r\n--", "text/html")[0].toString(), "<p>¿Sí?</p>");
});

test("buildMime is base64url of buildMimeMessage (Gmail)", () => {
  const raw = buildMime({ from: "a@x.com", to: ["b@x.com"], subject: "x", html: "<p>x</p>" });
  assert.doesNotMatch(raw, /[+/=]/);
  assert.match(Buffer.from(raw, "base64url").toString(), /^From: a@x\.com\r\nTo: b@x\.com/);
});

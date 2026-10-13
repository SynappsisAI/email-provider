import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { isExchangeInternalSender, MicrosoftEmailProvider } from "../src/providers/microsoft.js";

type Header = { name: string; value: string };

// Header sets shaped like Graph's `internetMessageHeaders` (anonymized).
const fixture = (name: string): Header[] =>
  JSON.parse(readFileSync(new URL(`./fixtures/${name}.headers.json`, import.meta.url), "utf8"));

const internal = fixture("internal");
const external = fixture("external");

const addr = (address: string) => ({ emailAddress: { address } });

// A Graph message as readMessage selects it: from + sender + headers.
const msg = (headers: Header[], from = "bob@contoso.com", sender: string | null = from) => ({
  from: addr(from),
  sender: sender === null ? undefined : addr(sender),
  internetMessageHeaders: headers,
});

const withHeader = (headers: Header[], name: string, value: string): Header[] =>
  headers.map((h) => (h.name === name ? { ...h, value } : h));

test("internal sender: all three stamps present once, sender = from → true", () => {
  assert.equal(isExchangeInternalSender(msg(internal)), true);
});

test("external sender → false", () => {
  assert.equal(isExchangeInternalSender(msg(external, "alice@example.net")), false);
});

test("forged from outside (Exchange overwrote the stamps, only the control header survived) → false", () => {
  const forged = fixture("forged-external");
  assert.ok(forged.some((h) => h.name === "X-Synappsis-Control"));
  assert.equal(isExchangeInternalSender(msg(forged, "bob@contoso.com")), false);
});

test("duplicated AuthAs → false, even when both copies say Internal", () => {
  const dupInternal = [...internal, { name: "X-MS-Exchange-Organization-AuthAs", value: "Internal" }];
  const dupMixed = [...internal, { name: "X-MS-Exchange-Organization-AuthAs", value: "Anonymous" }];
  assert.equal(isExchangeInternalSender(msg(dupInternal)), false);
  assert.equal(isExchangeInternalSender(msg(dupMixed)), false);
});

test("contradictory set (Internal + Incoming) → false", () => {
  const contradictory = withHeader(internal, "X-MS-Exchange-Organization-MessageDirectionality", "Incoming");
  assert.equal(isExchangeInternalSender(msg(contradictory)), false);
});

test("any single stamp with the external value → false", () => {
  for (const [name, value] of [
    ["X-MS-Exchange-Organization-AuthAs", "Anonymous"],
    ["X-MS-Exchange-Organization-MessageDirectionality", "Incoming"],
    ["X-MS-Exchange-CrossTenant-FromEntityHeader", "Internet"],
  ]) {
    assert.equal(isExchangeInternalSender(msg(withHeader(internal, name, value))), false, name);
  }
});

test("any stamp missing → false", () => {
  for (const name of [
    "X-MS-Exchange-Organization-AuthAs",
    "X-MS-Exchange-Organization-MessageDirectionality",
    "X-MS-Exchange-CrossTenant-FromEntityHeader",
  ]) {
    assert.equal(isExchangeInternalSender(msg(internal.filter((h) => h.name !== name))), false, name);
  }
});

test("sender ≠ from (send on behalf / delegate) → false", () => {
  assert.equal(isExchangeInternalSender(msg(internal, "boss@contoso.com", "assistant@contoso.com")), false);
});

test("sender or from missing → false", () => {
  assert.equal(isExchangeInternalSender(msg(internal, "bob@contoso.com", null)), false);
  assert.equal(isExchangeInternalSender({ internetMessageHeaders: internal }), false);
  assert.equal(isExchangeInternalSender(msg(internal, "", "")), false);
});

test("sender/from compare case-insensitively, trimmed", () => {
  assert.equal(isExchangeInternalSender(msg(internal, "Bob@Contoso.com", " bob@contoso.COM ")), true);
});

test("rule redirect / auto-forward / resent markers → false, even with internal stamps", () => {
  for (const [name, value] of [
    ["X-MS-Exchange-Organization-AutoForwarded", "true"],
    ["X-MS-Exchange-Inbox-Rules-Loop", "bob@contoso.com"],
    ["X-MS-Exchange-Generated-Message-Source", "Mailbox Rules Agent"],
    ["Resent-From", "bob@contoso.com"],
    ["RESENT-MESSAGE-ID", "<abc@contoso.com>"],
    ["Auto-Submitted", "auto-forwarded"],
    ["Auto-Submitted", "auto-replied"],
  ]) {
    assert.equal(isExchangeInternalSender(msg([...internal, { name, value }])), false, `${name}: ${value}`);
  }
});

test("Auto-Submitted: no is fine", () => {
  assert.equal(isExchangeInternalSender(msg([...internal, { name: "Auto-Submitted", value: " No " }])), true);
});

test("names and values compare case-insensitively, values trimmed", () => {
  const varied = internal.map((h) =>
    h.name.startsWith("X-MS-Exchange-") ? { name: h.name.toUpperCase(), value: `  ${h.value.toLowerCase()} ` } : h,
  );
  assert.equal(isExchangeInternalSender(msg(varied)), true);
});

test("no headers → false", () => {
  assert.equal(isExchangeInternalSender(msg([])), false);
  assert.equal(isExchangeInternalSender({ from: addr("bob@contoso.com"), sender: addr("bob@contoso.com") }), false);
});

test("toFull exposes internalSender as a strict boolean", () => {
  const toFull = (MicrosoftEmailProvider as any).toFull;
  assert.equal(toFull({ id: "1", ...msg(internal) }).internalSender, true);
  assert.equal(toFull({ id: "2", ...msg(external, "alice@example.net") }).internalSender, false);
  assert.equal(toFull({ id: "3" }).internalSender, false);
});

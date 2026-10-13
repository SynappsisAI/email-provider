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

const withHeader = (headers: Header[], name: string, value: string): Header[] =>
  headers.map((h) => (h.name === name ? { ...h, value } : h));

test("internal sender: all three stamps present once → true", () => {
  assert.equal(isExchangeInternalSender(internal), true);
});

test("external sender → false", () => {
  assert.equal(isExchangeInternalSender(external), false);
});

test("forged from outside (Exchange overwrote the stamps, only the control header survived) → false", () => {
  const forged = fixture("forged-external");
  assert.ok(forged.some((h) => h.name === "X-Synappsis-Control"));
  assert.equal(isExchangeInternalSender(forged), false);
});

test("duplicated AuthAs → false, even when both copies say Internal", () => {
  const dupInternal = [...internal, { name: "X-MS-Exchange-Organization-AuthAs", value: "Internal" }];
  const dupMixed = [...internal, { name: "X-MS-Exchange-Organization-AuthAs", value: "Anonymous" }];
  assert.equal(isExchangeInternalSender(dupInternal), false);
  assert.equal(isExchangeInternalSender(dupMixed), false);
});

test("contradictory set (Internal + Incoming) → false", () => {
  const contradictory = withHeader(internal, "X-MS-Exchange-Organization-MessageDirectionality", "Incoming");
  assert.equal(isExchangeInternalSender(contradictory), false);
});

test("any stamp missing → false", () => {
  for (const name of [
    "X-MS-Exchange-Organization-AuthAs",
    "X-MS-Exchange-Organization-MessageDirectionality",
    "X-MS-Exchange-CrossTenant-FromEntityHeader",
  ]) {
    assert.equal(isExchangeInternalSender(internal.filter((h) => h.name !== name)), false, name);
  }
});

test("names and values compare case-insensitively, values trimmed", () => {
  const varied = internal.map((h) =>
    h.name.startsWith("X-MS-Exchange-") ? { name: h.name.toUpperCase(), value: `  ${h.value.toLowerCase()} ` } : h,
  );
  assert.equal(isExchangeInternalSender(varied), true);
});

test("no headers → false", () => {
  assert.equal(isExchangeInternalSender(undefined), false);
  assert.equal(isExchangeInternalSender([]), false);
});

test("toFull exposes internalSender from internetMessageHeaders", () => {
  const toFull = (MicrosoftEmailProvider as any).toFull;
  assert.equal(toFull({ id: "1", internetMessageHeaders: internal }).internalSender, true);
  assert.equal(toFull({ id: "2", internetMessageHeaders: external }).internalSender, false);
  assert.equal(toFull({ id: "3" }).internalSender, false);
});

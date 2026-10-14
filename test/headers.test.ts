import { test } from "node:test";
import assert from "node:assert/strict";
import { assertOriginHeaderName, classifyKind, originHeaderPair, readOrigin } from "../src/headers.js";

const H = (...pairs: [string, string][]) => pairs.map(([name, value]) => ({ name, value }));

test("classifyKind: plain message", () => {
  assert.equal(classifyKind({ headers: H(["Content-Type", "text/html"]), fromAddress: "bob@contoso.com" }), "message");
});

test("classifyKind: NDR by class, report type or postmaster sender", () => {
  assert.equal(classifyKind({ headers: [], messageClass: "REPORT.IPM.Note.NDR" }), "ndr");
  assert.equal(classifyKind({ headers: H(["Content-Type", "multipart/report; report-type=delivery-status"]) }), "ndr");
  assert.equal(classifyKind({ headers: [], fromAddress: "MAILER-DAEMON@example.net" }), "ndr");
});

test("classifyKind: read receipt by class or disposition-notification report", () => {
  assert.equal(classifyKind({ headers: [], messageClass: "REPORT.IPM.Note.IPNRN" }), "read-receipt");
  assert.equal(classifyKind({ headers: [], messageClass: "REPORT.IPM.Note.IPNNRN" }), "read-receipt");
  assert.equal(
    classifyKind({ headers: H(["Content-Type", 'multipart/report; report-type="disposition-notification"']) }),
    "read-receipt",
  );
});

test("classifyKind: meeting by class, eventMessage or calendar part", () => {
  assert.equal(classifyKind({ headers: [], messageClass: "IPM.Schedule.Meeting.Resp.Pos" }), "meeting");
  assert.equal(classifyKind({ headers: [], isEventMessage: true }), "meeting");
  assert.equal(classifyKind({ headers: [], hasCalendarPart: true }), "meeting");
});

test("classifyKind: auto-reply by OOF class or headers; Auto-Submitted: no is a message", () => {
  assert.equal(classifyKind({ headers: [], messageClass: "IPM.Note.Rules.OofTemplate.Microsoft" }), "auto-reply");
  assert.equal(classifyKind({ headers: H(["Auto-Submitted", "auto-replied"]) }), "auto-reply");
  assert.equal(classifyKind({ headers: H(["Precedence", "auto_reply"]) }), "auto-reply");
  assert.equal(classifyKind({ headers: H(["Auto-Submitted", "no"]) }), "message");
});

test("classifyKind: X-Auto-Response-Suppress alone is a message (Exchange stamps it on distribution-group mail)", () => {
  assert.equal(classifyKind({ headers: H(["X-Auto-Response-Suppress", "All"]) }), "message");
});

test("classifyKind precedence: an auto-generated NDR is still ndr", () => {
  assert.equal(classifyKind({ headers: H(["Auto-Submitted", "auto-replied"]), messageClass: "REPORT.IPM.Note.NDR" }), "ndr");
});

test("origin header name must be X-…", () => {
  assert.equal(assertOriginHeaderName("X-AgentFleet-Origin"), "X-AgentFleet-Origin");
  for (const bad of ["Origin", "X-Bad Name", "X-Evil\r\nBcc", ""]) assert.throws(() => assertOriginHeaderName(bad), bad);
});

test("originHeaderPair: none asked → []; asked without config → throws; bad value → throws", () => {
  assert.deepEqual(originHeaderPair(undefined, undefined), []);
  assert.deepEqual(originHeaderPair("X-O", " agentfleet; kind=digest "), [["X-O", "agentfleet; kind=digest"]]);
  assert.throws(() => originHeaderPair(undefined, "agentfleet"), /originHeader/);
  for (const bad of ["", "a\r\nBcc: x", "ñ", "x".repeat(201)]) assert.throws(() => originHeaderPair("X-O", bad));
});

test("readOrigin: configured header, case-insensitive; undefined when unconfigured or absent", () => {
  const headers = H(["x-agentfleet-origin", " agentfleet "]);
  assert.equal(readOrigin("X-AgentFleet-Origin", headers), "agentfleet");
  assert.equal(readOrigin(undefined, headers), undefined);
  assert.equal(readOrigin("X-AgentFleet-Origin", H(["Subject", "x"])), undefined);
});

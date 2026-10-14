import type { MessageKind } from "./types.js";

// Header helpers shared by both providers: origin-header validation, header-value
// sanitizing, and the header-based `MessageFull.kind` classifier.

export type HeaderList = { name?: string | null; value?: string | null }[] | null | undefined;

/** Value of the first header named `name` (case-insensitive), trimmed; undefined if absent. */
export function findHeader(headers: HeaderList, name: string): string | undefined {
  const lower = name.toLowerCase();
  const h = (headers ?? []).find((h) => (h?.name ?? "").trim().toLowerCase() === lower);
  return h ? (h.value ?? "").trim() : undefined;
}

/** Drop CR/LF so a caller-supplied value can never start a new header (header injection). */
export function headerValue(v: string): string {
  return v.replace(/[\r\n]+/g, " ").trim();
}

/** Graph only accepts custom `internetMessageHeaders` named `x-…`; Gmail keeps any. One rule for both. */
export function assertOriginHeaderName(name: string): string {
  if (!/^X-[A-Za-z0-9-]+$/i.test(name)) {
    throw new Error(`originHeader must be an X- header name (letters, digits, '-'), got "${name}"`);
  }
  return name;
}

/** The [name, value] pair to stamp on a send, or [] when no origin was asked for. */
export function originHeaderPair(configured: string | undefined, origin: string | undefined): [string, string][] {
  if (origin === undefined) return [];
  if (!configured) throw new Error("origin requires originHeader in the provider config");
  const value = origin.trim();
  if (!/^[\x21-\x7E][\x20-\x7E]{0,199}$/.test(value)) {
    throw new Error("origin must be 1–200 printable ASCII characters");
  }
  return [[configured, value]];
}

/** The origin value read back from a message, or undefined (not configured / absent). */
export function readOrigin(configured: string | undefined, headers: HeaderList): string | undefined {
  return configured ? findHeader(headers, configured) : undefined;
}

/**
 * Header-based classification shared by both providers. Provider-specific signals
 * come in as extras: Exchange's message class (`PR_MESSAGE_CLASS`), Graph's
 * eventMessage type, a text/calendar part. Precedence: ndr > read-receipt >
 * meeting > auto-reply > message.
 *
 * `X-Auto-Response-Suppress` is deliberately NOT a signal: it asks the receiver not to
 * auto-respond, and Exchange stamps it on ordinary mail delivered through distribution
 * groups — treating it as automatic would drop real human messages.
 */
export function classifyKind(input: {
  headers: HeaderList;
  fromAddress?: string;
  messageClass?: string;
  isEventMessage?: boolean;
  hasCalendarPart?: boolean;
}): MessageKind {
  const cls = (input.messageClass ?? "").toLowerCase();
  const contentType = (findHeader(input.headers, "Content-Type") ?? "").toLowerCase();
  const isReport = contentType.startsWith("multipart/report");
  const reportType = contentType.match(/report-type\s*=\s*"?([a-z-]+)/)?.[1];
  const fromLocal = (input.fromAddress ?? "").toLowerCase().split("@")[0];

  if (cls.startsWith("report.") && cls.endsWith(".ndr")) return "ndr";
  if (isReport && reportType !== "disposition-notification") return "ndr";
  if (fromLocal === "postmaster" || fromLocal === "mailer-daemon") return "ndr";

  if (cls.startsWith("report.") && /\.ipn(n)?rn$/.test(cls)) return "read-receipt";
  if (isReport && reportType === "disposition-notification") return "read-receipt";

  if (cls.startsWith("ipm.schedule.meeting.") || input.isEventMessage || input.hasCalendarPart) return "meeting";

  const autoSubmitted = findHeader(input.headers, "Auto-Submitted")?.toLowerCase();
  const precedence = findHeader(input.headers, "Precedence")?.toLowerCase();
  if (
    cls.startsWith("ipm.note.rules.ooftemplate") ||
    cls.startsWith("ipm.note.rules.replytemplate") ||
    (autoSubmitted !== undefined && autoSubmitted !== "no") ||
    precedence === "auto_reply"
  ) {
    return "auto-reply";
  }
  return "message";
}

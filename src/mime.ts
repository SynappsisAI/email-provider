import type { Attachment } from "./types.js";
import { headerValue } from "./headers.js";

/** `filename` parameter: quoted ASCII, or RFC 2231 `filename*` for non-ASCII / quotes /
 *  control chars, which would otherwise break the header. */
function filenameParam(name: string): string {
  if (/^[\x20-\x7E]*$/.test(name) && !/["\\]/.test(name)) return `filename="${name}"`;
  const encoded = encodeURIComponent(name.replace(/[\r\n]/g, " ")).replace(/['()*]/g, (c) => `%${c.charCodeAt(0).toString(16).toUpperCase()}`);
  return `filename*=UTF-8''${encoded}`;
}

/** RFC 2047-encode a header value that has non-ASCII characters. */
function encodeWord(v: string): string {
  return /[^\x00-\x7F]/.test(v) ? `=?UTF-8?B?${Buffer.from(v, "utf-8").toString("base64")}?=` : v;
}

/** base64 wrapped at 76 chars per RFC 2045, safe for picky clients and long lines. */
function base64Lines(buf: Buffer): string {
  return buf.toString("base64").replace(/(.{76})/g, "$1\r\n");
}

export interface MimeOptions {
  from: string;
  to: string[];
  cc?: string[];
  bcc?: string[];
  subject: string;
  html: string;
  /** Plain-text alternative: the body becomes multipart/alternative (text, then html). */
  text?: string;
  /** Our own `Message-ID`, with angle brackets. */
  messageId?: string;
  inReplyTo?: string;
  references?: string;
  /** Extra headers, e.g. the origin header or `Auto-Submitted`. */
  headers?: [string, string][];
  attachments?: Attachment[];
}

/**
 * Build an RFC 5322/2045 MIME message as a string. Every caller-supplied header value is
 * stripped of CR/LF (no header injection); text parts and attachments are base64 so long
 * lines and non-ASCII bodies survive any relay.
 *
 * Attachments come in as `Buffer` and are base64-encoded inside the MIME body (the wire
 * format MIME requires). The library never asks the caller to pre-encode anything.
 *
 *   mixed                     (only with regular attachments)
 *   ├── alternative           (only with `text`)
 *   │   ├── text/plain
 *   │   └── related           (only with inline parts) ── html + inline images
 *   └── attachments…
 */
export function buildMimeMessage(opts: MimeOptions): string {
  const hv = headerValue;
  const lines: string[] = [`From: ${hv(opts.from)}`, `To: ${opts.to.map(hv).join(", ")}`];
  if (opts.cc?.length) lines.push(`Cc: ${opts.cc.map(hv).join(", ")}`);
  if (opts.bcc?.length) lines.push(`Bcc: ${opts.bcc.map(hv).join(", ")}`);
  lines.push(`Subject: ${encodeWord(hv(opts.subject))}`);
  if (opts.messageId) lines.push(`Message-ID: ${hv(opts.messageId)}`);
  if (opts.inReplyTo) {
    lines.push(`In-Reply-To: ${hv(opts.inReplyTo)}`);
    lines.push(`References: ${hv(opts.references ?? opts.inReplyTo)}`);
  } else if (opts.references) {
    lines.push(`References: ${hv(opts.references)}`);
  }
  for (const [name, value] of opts.headers ?? []) lines.push(`${hv(name)}: ${encodeWord(hv(value))}`);
  lines.push("MIME-Version: 1.0");

  const inline = (opts.attachments ?? []).filter((a) => a.contentId);
  const regular = (opts.attachments ?? []).filter((a) => !a.contentId);
  const newBoundary = () => `synappsis_${Date.now()}_${Math.random().toString(36).slice(2)}`;

  const pushText = (type: "html" | "plain", content: string) => {
    lines.push(`Content-Type: text/${type}; charset="UTF-8"`);
    lines.push("Content-Transfer-Encoding: base64");
    lines.push("");
    lines.push(base64Lines(Buffer.from(content, "utf-8")));
  };
  const pushPart = (att: Attachment) => {
    lines.push(`Content-Type: ${hv(att.contentType)}`);
    lines.push("Content-Transfer-Encoding: base64");
    if (att.contentId) {
      lines.push(`Content-ID: <${hv(att.contentId)}>`);
      lines.push(`Content-Disposition: inline; ${filenameParam(att.filename)}`);
    } else {
      lines.push(`Content-Disposition: attachment; ${filenameParam(att.filename)}`);
    }
    lines.push("");
    lines.push(base64Lines(att.content));
  };
  const multipart = (type: string, children: (() => void)[], extra = "") => {
    const boundary = newBoundary();
    lines.push(`Content-Type: multipart/${type};${extra} boundary="${boundary}"`);
    lines.push("");
    for (const child of children) {
      lines.push(`--${boundary}`);
      child();
    }
    lines.push(`--${boundary}--`);
  };

  // The HTML, wrapped with its inline parts in multipart/related if any.
  const html = () =>
    inline.length
      ? multipart("related", [() => pushText("html", opts.html), ...inline.map((a) => () => pushPart(a))], ' type="text/html";')
      : pushText("html", opts.html);
  const body = () => (opts.text !== undefined ? multipart("alternative", [() => pushText("plain", opts.text!), html]) : html());

  if (regular.length) multipart("mixed", [body, ...regular.map((a) => () => pushPart(a))]);
  else body();

  return lines.join("\r\n");
}

/** `buildMimeMessage`, base64url-encoded for the Gmail API's `users.messages.send`. */
export function buildMime(opts: MimeOptions): string {
  return Buffer.from(buildMimeMessage(opts)).toString("base64url");
}

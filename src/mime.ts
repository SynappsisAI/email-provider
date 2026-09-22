import type { Attachment } from "./types.js";

/**
 * Build a base64url-encoded RFC 2822/2045 MIME message ready for the
 * Gmail API's `users.messages.send`. Used by the Google provider only —
 * Microsoft Graph takes structured JSON instead of raw MIME.
 *
 * Attachments come in as `Buffer` and are base64-encoded inside the MIME
 * body (which is the wire format MIME requires). The library never asks
 * the caller to pre-encode anything.
 */
export function buildMime(opts: {
  from: string;
  to: string[];
  cc?: string[];
  bcc?: string[];
  subject: string;
  html: string;
  inReplyTo?: string;
  references?: string;
  attachments?: Attachment[];
}): string {
  const lines: string[] = [
    `From: ${opts.from}`,
    `To: ${opts.to.join(", ")}`,
  ];
  if (opts.cc?.length) lines.push(`Cc: ${opts.cc.join(", ")}`);
  if (opts.bcc?.length) lines.push(`Bcc: ${opts.bcc.join(", ")}`);

  // RFC 2047 encode the subject if it contains non-ASCII characters.
  const hasNonAscii = /[^\x00-\x7F]/.test(opts.subject);
  const encodedSubject = hasNonAscii
    ? `=?UTF-8?B?${Buffer.from(opts.subject, "utf-8").toString("base64")}?=`
    : opts.subject;
  lines.push(`Subject: ${encodedSubject}`);

  if (opts.inReplyTo) {
    lines.push(`In-Reply-To: ${opts.inReplyTo}`);
    lines.push(`References: ${opts.references ?? opts.inReplyTo}`);
  }
  lines.push("MIME-Version: 1.0");

  // Structure: inline (cid) parts go in a multipart/related next to the HTML;
  // regular attachments wrap that in a multipart/mixed.
  //   mixed
  //   ├── related (only if inline parts) ── html + inline images
  //   └── attachments…
  const inline = (opts.attachments ?? []).filter((a) => a.contentId);
  const regular = (opts.attachments ?? []).filter((a) => !a.contentId);
  const newBoundary = () => `synappsis_${Date.now()}_${Math.random().toString(36).slice(2)}`;

  const pushHtml = () => {
    lines.push('Content-Type: text/html; charset="UTF-8"');
    lines.push("");
    lines.push(opts.html);
  };
  const pushPart = (att: Attachment) => {
    lines.push(`Content-Type: ${att.contentType}`);
    lines.push("Content-Transfer-Encoding: base64");
    if (att.contentId) {
      lines.push(`Content-ID: <${att.contentId}>`);
      lines.push(`Content-Disposition: inline; filename="${att.filename}"`);
    } else {
      lines.push(`Content-Disposition: attachment; filename="${att.filename}"`);
    }
    lines.push("");
    // base64-encode the raw buffer for MIME wire format.
    // Wrap at 76 chars per RFC 2045 to be safe with picky clients.
    lines.push(att.content.toString("base64").replace(/(.{76})/g, "$1\r\n"));
  };
  // Body = the HTML, wrapped with its inline parts in multipart/related if any.
  const pushBody = () => {
    if (!inline.length) return pushHtml();
    const related = newBoundary();
    lines.push(`Content-Type: multipart/related; boundary="${related}"`);
    lines.push("");
    lines.push(`--${related}`);
    pushHtml();
    for (const att of inline) {
      lines.push(`--${related}`);
      pushPart(att);
    }
    lines.push(`--${related}--`);
  };

  if (regular.length) {
    const mixed = newBoundary();
    lines.push(`Content-Type: multipart/mixed; boundary="${mixed}"`);
    lines.push("");
    lines.push(`--${mixed}`);
    pushBody();
    for (const att of regular) {
      lines.push(`--${mixed}`);
      pushPart(att);
    }
    lines.push(`--${mixed}--`);
  } else {
    pushBody();
  }

  // Gmail API wants base64url (not standard base64) of the entire MIME bytes.
  return Buffer.from(lines.join("\r\n"))
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
}

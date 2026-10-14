// Keyset cursor for `listSince`: the receive time of the last message returned plus the
// ids already returned AT that time. The next page asks again for everything received
// at or after that time and skips those ids, so a message marked read or moved mid-walk
// can't shift a later one past the page boundary — unlike offset paging ($skip / slice).

export interface Cursor {
  /** Provider tag, so a token from one provider can't be replayed on the other. */
  p: "microsoft" | "google";
  /** Receive time of the last message returned, epoch ms. */
  t: number;
  /** Ids already returned with exactly that receive time. */
  ids: string[];
}

export function encodeCursor(c: Cursor): string {
  return Buffer.from(JSON.stringify(c)).toString("base64url");
}

export function decodeCursor(token: string, provider: Cursor["p"]): Cursor {
  let c: any;
  try {
    c = JSON.parse(Buffer.from(token, "base64url").toString("utf-8"));
  } catch {
    throw new Error("invalid pageToken");
  }
  const ok =
    c?.p === provider &&
    Number.isSafeInteger(c.t) &&
    Array.isArray(c.ids) &&
    c.ids.every((id: unknown) => typeof id === "string");
  if (!ok) throw new Error("invalid pageToken");
  return c as Cursor;
}

/** The cursor after a page: its last item's time, plus every returned id at that time
 *  (and the previous cursor's ids too when the time didn't move). */
export function nextCursor<T>(
  p: Cursor["p"],
  page: T[],
  time: (x: T) => number,
  id: (x: T) => string,
  prev: Cursor | undefined,
): Cursor {
  if (!page.length) throw new Error("nextCursor: empty page");
  const t = time(page[page.length - 1]);
  const ids = page.filter((x) => time(x) === t).map(id);
  if (prev && prev.t === t) ids.unshift(...prev.ids);
  return { p, t, ids };
}

/** listSince page size: default 25, max 100; anything else is a caller bug. */
export function pageSizeOf(limit: number | undefined): number {
  if (limit === undefined) return 25;
  if (!Number.isInteger(limit) || limit < 1) throw new Error(`limit must be an integer >= 1, got ${limit}`);
  return Math.min(limit, 100);
}

/** `since` as epoch ms; throws on anything Date can't parse. */
export function parseSince(since: string): number {
  const ms = Date.parse(since);
  if (Number.isNaN(ms)) throw new Error(`since must be an ISO timestamp, got "${since}"`);
  return ms;
}

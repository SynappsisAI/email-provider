import { test, beforeEach, afterEach } from "node:test";
import assert from "node:assert/strict";
import { MicrosoftEmailProvider } from "../src/providers/microsoft.js";
import { GoogleEmailProvider } from "../src/providers/google.js";
import type { EmailProvider, ListSinceParams, MessageSummary } from "../src/types.js";

// In-memory mailboxes that behave like each API on the points listSince depends on, and
// that can change between page calls (marked read, moved, deleted).

type Msg = { id: string; ms: number; unread: boolean; folder: string; deleted?: boolean };

/** Deterministic pseudo-random shuffle: each call orders ties (and Gmail lists) differently. */
function shuffle<T>(xs: T[], seed: number): T[] {
  const a = [...xs];
  let s = seed || 1;
  for (let i = a.length - 1; i > 0; i--) {
    s = (s * 1103515245 + 12345) % 2147483648;
    const j = s % (i + 1);
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

const BASE = Math.floor(Date.now() / 1000 - 86400) * 1000; // a day ago, whole second

// ── Microsoft (Graph) ──

let realFetch: typeof fetch;
let graphCalls: { url: URL; headers: Record<string, string> }[] = [];
beforeEach(() => {
  realFetch = globalThis.fetch;
  graphCalls = [];
});
afterEach(() => {
  globalThis.fetch = realFetch;
});

/** Graph: receivedDateTime has whole-second precision; ties come back in a different order per call. */
function fakeGraph(box: Msg[]) {
  let call = 0;
  globalThis.fetch = (async (input: string, init: RequestInit = {}) => {
    const url = new URL(input);
    graphCalls.push({ url, headers: init.headers as Record<string, string> });
    const folder = url.pathname.match(/mailFolders\/([^/]+)\/messages$/)![1];
    const filter = url.searchParams.get("$filter")!;
    const ge = Date.parse(filter.match(/receivedDateTime ge (\S+)/)![1]);
    const unreadOnly = /isRead eq false/.test(filter);
    assert.equal(url.searchParams.get("$orderby"), "receivedDateTime asc");
    const top = Number(url.searchParams.get("$top"));
    const sec = (m: Msg) => Math.floor(m.ms / 1000) * 1000;
    const matching = box.filter((m) => !m.deleted && m.folder === folder && m.ms >= ge && (!unreadOnly || m.unread));
    const ordered = shuffle(matching, ++call).sort((a, b) => sec(a) - sec(b));
    const value = ordered.slice(0, top).map((m) => ({
      id: m.id,
      subject: m.id,
      receivedDateTime: new Date(sec(m)).toISOString().replace(".000Z", "Z"),
      isRead: !m.unread,
    }));
    const body: any = { value };
    if (ordered.length > top) body["@odata.nextLink"] = `${url.origin}${url.pathname}?$skip=${top}`;
    return new Response(JSON.stringify(body), { status: 200 });
  }) as typeof fetch;
}

function microsoft(extra: { immutableIds?: boolean } = {}) {
  const p = new MicrosoftEmailProvider({ type: "microsoft", clientId: "c", clientSecret: "s", tenantId: "t", ...extra });
  (p as any).getToken = async () => "token";
  return p;
}

// ── Google (Gmail) ──

/**
 * Gmail: ids only, in no particular order, ≤500 per page with a pageToken. `after:`/`before:`
 * are in seconds; their exact inclusiveness isn't documented, so `inclusive` toggles the two
 * plausible readings and every Gmail test runs under both.
 */
function google(box: Msg[], inclusive: boolean) {
  let listCall = 0;
  let inFlight = 0;
  const stats = { maxInFlight: 0, listCalls: 0 };
  const err404 = () => Object.assign(new Error("Not Found"), { code: 404 });
  const fake = {
    users: {
      messages: {
        list: async (a: any) => {
          stats.listCalls++;
          const after = Number(a.q.match(/after:(-?\d+)/)[1]);
          const before = Number(a.q.match(/before:(-?\d+)/)[1]);
          const unread = /is:unread/.test(a.q);
          const inWindow = (ms: number) =>
            inclusive ? ms >= after * 1000 && ms < (before + 1) * 1000 : ms >= (after + 1) * 1000 && ms < before * 1000;
          const ids = shuffle(
            box.filter((m) => !m.deleted && a.labelIds.includes(m.folder) && inWindow(m.ms) && (!unread || m.unread)),
            ++listCall,
          ).map((m) => ({ id: m.id }));
          const start = Number(a.pageToken ?? 0);
          const size = Math.min(a.maxResults, 500);
          const page = ids.slice(start, start + size);
          return { data: { messages: page, nextPageToken: start + size < ids.length ? String(start + size) : undefined } };
        },
        get: async (a: any) => {
          inFlight++;
          stats.maxInFlight = Math.max(stats.maxInFlight, inFlight);
          await new Promise((r) => setImmediate(r));
          inFlight--;
          const m = box.find((x) => x.id === a.id);
          if (!m || m.deleted) throw err404();
          return { data: { id: m.id, threadId: m.id, internalDate: String(m.ms), payload: { headers: [{ name: "Subject", value: m.id }] } } };
        },
      },
    },
  };
  const p = new GoogleEmailProvider({ type: "google", serviceAccountKey: {} });
  (p as any).getGmail = () => fake;
  return { p, stats };
}

// ── Shared walk + assertions ──

async function walk(
  p: EmailProvider,
  params: Omit<ListSinceParams, "pageToken">,
  between: (pageIndex: number, returned: MessageSummary[]) => void = () => {},
) {
  const out: MessageSummary[] = [];
  let pageToken: string | undefined;
  for (let i = 0; i < 1000; i++) {
    const r = await p.listSince({ ...params, pageToken });
    out.push(...r.messages);
    if (!r.nextPageToken) return out;
    pageToken = r.nextPageToken;
    between(i, out);
  }
  throw new Error("walk did not end");
}

function assertAscendingOnce(got: MessageSummary[], expectedIds: string[]) {
  const ids = got.map((m) => m.id);
  assert.equal(new Set(ids).size, ids.length, "no duplicates");
  assert.deepEqual([...ids].sort(), [...expectedIds].sort(), "exactly the expected messages");
  for (let i = 1; i < got.length; i++) {
    assert.ok(Date.parse(got[i - 1].receivedAt) <= Date.parse(got[i].receivedAt), `ascending at ${i}`);
  }
}

/** 1200 messages; Gmail times get ms jitter, Graph rounds to seconds → many ties. */
function bigBox(n = 1200): Msg[] {
  return Array.from({ length: n }, (_, i) => ({
    id: `m${String(i).padStart(5, "0")}`,
    ms: BASE + Math.floor(i / 3) * 1000 + (i % 3) * 7, // 3 per second
    unread: true,
    folder: "inbox",
  }));
}

const providers = [
  { name: "microsoft", make: (box: Msg[]) => (fakeGraph(box), microsoft()), folder: "inbox" },
  { name: "google (exclusive bounds)", make: (box: Msg[]) => google(box, false).p, folder: "INBOX" },
  { name: "google (inclusive bounds)", make: (box: Msg[]) => google(box, true).p, folder: "INBOX" },
];

for (const prov of providers) {
  const boxFor = (n?: number) => bigBox(n).map((m) => ({ ...m, folder: prov.folder }));

  test(`${prov.name}: 1200 messages since the timestamp, each exactly once, ascending`, async () => {
    const box = boxFor();
    const got = await walk(prov.make(box), { mailbox: "a@x.com", since: new Date(BASE).toISOString(), limit: 100 });
    assertAscendingOnce(got, box.map((m) => m.id));
  });

  test(`${prov.name}: marking read / moving / deleting returned messages mid-walk skips nothing`, async () => {
    const box = boxFor();
    // Some not-yet-returned messages also get read mid-walk: those may legitimately drop out.
    const volatile = new Set<string>();
    const got = await walk(prov.make(box), { mailbox: "a@x.com", since: new Date(BASE).toISOString(), unreadOnly: true, limit: 50 }, (i, returned) => {
      // The offset-paging killer: what was just returned stops matching the query.
      for (const m of returned.slice(-50)) {
        const msg = box.find((x) => x.id === m.id)!;
        if (i % 3 === 0) msg.unread = false;
        else if (i % 3 === 1) msg.folder = "archive";
        else msg.deleted = true;
      }
      const ahead = box.find((x) => x.unread && !x.deleted && x.folder === prov.folder && !returned.some((r) => r.id === x.id));
      if (ahead && i % 5 === 0) {
        ahead.unread = false;
        volatile.add(ahead.id);
      }
    });
    const mustSee = box.map((m) => m.id).filter((id) => !volatile.has(id));
    const ids = new Set(got.map((m) => m.id));
    assert.equal(ids.size, got.length, "no duplicates");
    for (const id of mustSee) assert.ok(ids.has(id), `missing ${id}`);
    for (let i = 1; i < got.length; i++) assert.ok(Date.parse(got[i - 1].receivedAt) <= Date.parse(got[i].receivedAt));
  });

  test(`${prov.name}: 250 messages with the same receive time across page boundaries`, async () => {
    const box = Array.from({ length: 250 }, (_, i) => ({ id: `t${i}`, ms: BASE + 5000, unread: true, folder: prov.folder }));
    box.push({ id: "before", ms: BASE + 4000, unread: true, folder: prov.folder }, { id: "after", ms: BASE + 6000, unread: true, folder: prov.folder });
    const got = await walk(prov.make(box), { mailbox: "a@x.com", since: new Date(BASE).toISOString(), limit: 100 });
    assertAscendingOnce(got, box.map((m) => m.id));
    assert.equal(got[0].id, "before");
    assert.equal(got[got.length - 1].id, "after");
  });

  test(`${prov.name}: since is inclusive; earlier messages are excluded`, async () => {
    const since = BASE + 10_000;
    const box = [
      { id: "older", ms: since - 1000, unread: true, folder: prov.folder },
      { id: "exact", ms: since, unread: true, folder: prov.folder },
      { id: "newer", ms: since + 1000, unread: true, folder: prov.folder },
    ];
    const got = await walk(prov.make(box), { mailbox: "a@x.com", since: new Date(since).toISOString() });
    assert.deepEqual(got.map((m) => m.id), ["exact", "newer"]);
  });

  test(`${prov.name}: nothing since → empty page, no token`, async () => {
    const box = boxFor(5);
    const r = await prov.make(box).listSince({ mailbox: "a@x.com", since: new Date(BASE + 3_600_000).toISOString() });
    assert.deepEqual(r, { messages: [] });
  });
}

test("google: after: second-granularity edge — exact internalDate bounds within one second", async () => {
  for (const inclusive of [false, true]) {
    const since = BASE + 20_500; // mid-second
    const box = [
      { id: "same-second-before", ms: since - 1, unread: true, folder: "INBOX" },
      { id: "exact", ms: since, unread: true, folder: "INBOX" },
      { id: "same-second-after", ms: since + 499, unread: true, folder: "INBOX" },
    ];
    const got = await walk(google(box, inclusive).p, { mailbox: "a@x.com", since: new Date(since).toISOString() });
    assert.deepEqual(got.map((m) => m.id), ["exact", "same-second-after"], `inclusive=${inclusive}`);
  }
});

test("google: messages.get is throttled (≤10 in flight)", async () => {
  const box = bigBox(300).map((m) => ({ ...m, folder: "INBOX" }));
  const { p, stats } = google(box, false);
  await walk(p, { mailbox: "a@x.com", since: new Date(BASE).toISOString(), limit: 100 });
  assert.ok(stats.maxInFlight <= 10, `max in flight ${stats.maxInFlight}`);
});

test("google: a burst near the start of a long range is found in O(log range) list calls", async () => {
  for (const inclusive of [false, true]) {
    const since = Math.floor(Date.now() / 1000 - 180 * 86400) * 1000; // 180 days back
    const burstAt = since + 1.3 * 3_600_000;
    const box = Array.from({ length: 60 }, (_, i) => ({ id: `b${i}`, ms: burstAt + i * 10, unread: true, folder: "INBOX" }));
    const { p, stats } = google(box, inclusive);
    const got = await walk(p, { mailbox: "a@x.com", since: new Date(since).toISOString(), limit: 25 });
    assertAscendingOnce(got, box.map((m) => m.id));
    assert.ok(stats.listCalls < 150, `inclusive=${inclusive}: ${stats.listCalls} list calls for 3 pages`);
  }
});

test("microsoft: a short page with a nextLink but only seen ids keeps the cursor instead of crashing", async () => {
  const T = new Date(BASE).toISOString().replace(".000Z", "Z");
  let call = 0;
  globalThis.fetch = (async () => {
    call++;
    const value = (call === 1 ? ["a", "b", "c"] : ["a", "b"]).map((id) => ({ id, receivedDateTime: T, isRead: false }));
    return new Response(JSON.stringify({ value, "@odata.nextLink": "https://graph.microsoft.com/v1.0/x?$skip=2" }), { status: 200 });
  }) as typeof fetch;
  const p = microsoft();
  const r1 = await p.listSince({ mailbox: "a@x.com", since: T, limit: 2 });
  assert.deepEqual(r1.messages.map((m) => m.id), ["a", "b"]);
  const r2 = await p.listSince({ mailbox: "a@x.com", since: T, limit: 2, pageToken: r1.nextPageToken });
  assert.deepEqual(r2, { messages: [], nextPageToken: r1.nextPageToken });
});

test("limit must be a positive integer", async () => {
  fakeGraph([]);
  for (const limit of [0, -1, 2.5]) {
    await assert.rejects(microsoft().listSince({ mailbox: "a@x.com", since: new Date(BASE).toISOString(), limit }), /limit must be/);
    await assert.rejects(google([], false).p.listSince({ mailbox: "a@x.com", since: new Date(BASE).toISOString(), limit }), /limit must be/);
  }
});

test("microsoft: honors immutableIds and filters unread with receivedDateTime first", async () => {
  fakeGraph([]);
  await microsoft({ immutableIds: true }).listSince({ mailbox: "a@x.com", since: new Date(BASE).toISOString(), unreadOnly: true });
  assert.equal(graphCalls[0].headers.Prefer, 'IdType="ImmutableId"');
  assert.match(graphCalls[0].url.searchParams.get("$filter")!, /^receivedDateTime ge \S+ and isRead eq false$/);
});

test("pageToken is opaque and provider-bound; garbage and bad since are rejected", async () => {
  fakeGraph(bigBox(30));
  const ms = microsoft();
  const r = await ms.listSince({ mailbox: "a@x.com", since: new Date(BASE).toISOString(), limit: 10 });
  assert.ok(r.nextPageToken);
  const g = google([], false).p;
  await assert.rejects(g.listSince({ mailbox: "a@x.com", since: new Date(BASE).toISOString(), pageToken: r.nextPageToken }), /invalid pageToken/);
  await assert.rejects(ms.listSince({ mailbox: "a@x.com", since: "x", pageToken: "not-a-token" }), /invalid pageToken/);
  await assert.rejects(ms.listSince({ mailbox: "a@x.com", since: "yesterday" }), /since must be an ISO timestamp/);
});

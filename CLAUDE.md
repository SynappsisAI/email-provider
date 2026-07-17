# CLAUDE.md — @synappsis/email-provider

## What this repo is

A **provider-agnostic email library** for Synappsis services. It wraps Google
Workspace (Gmail API) and Microsoft 365 (Graph API) behind a single
`EmailProvider` interface, so consumers send/read/search mail without caring
which backend a mailbox lives on.

It is deliberately a **pure-mechanics** library: send, list, read, search, reply,
forward, get attachments, mark read, list folders — plus MIME construction and
Buffer-based attachments (no base64 in the public API). It does **not** contain an
HTTP server, MCP transport, authorization, policy, API-key validation, or
per-caller rate limiting. Those belong to the consumers:

- **synmail** — the HTTP server that fronts this library with API-key whitelists.
- **AgentFleet** — wires this library into per-tenant channel config; bundles the
  source via esbuild in a CDK `NodejsFunction`.

## Read first

- `handoff-history/` — session handoffs, newest = highest `HANDOFF-<N>.md`. Run
  `/read-handoff` at the start of a session to resume prior work. (Created on the
  first `/handoff`; may not exist yet.)
- `README.md` — the public contract, usage examples, provider setup (Google
  domain-wide delegation scopes, Microsoft app-permission grants), and the
  no-build distribution model. Keep it in sync with `src/types.ts`.
- The auto-memory dir for this repo (Claude's per-project memory) for accumulated
  context and prior decisions.

## Architecture map

```
src/
  index.ts              Public surface — the only entry point consumers import.
  types.ts              Normalized cross-provider types + the EmailProvider contract.
  mime.ts               Gmail-only: builds base64url RFC 2822 MIME (Graph takes JSON).
  providers/
    factory.ts          createEmailProvider(config) — switches on config.type.
    google.ts           GoogleEmailProvider — Gmail API + domain-wide delegation.
    microsoft.ts        MicrosoftEmailProvider — Graph API + tenant app permissions.
```

The two providers implement the same `EmailProvider` interface and normalize into
the shared types in `types.ts`. Encoding differences are hidden internally: Google
base64-encodes buffers inside the MIME body; Microsoft base64-encodes for Graph's
`contentBytes`. **Callers only ever pass and receive `Buffer`.**

## Dev workflow

- **Type-check:** `npm run typecheck` (`tsc --noEmit`, strict mode). This is the
  only gate — there is no build and no test suite yet. Type-check before every
  commit.
- **No build step.** `main` and `types` both point at `src/index.ts`; consumers
  import the TypeScript source directly (synmail via `tsx`, AgentFleet via esbuild).
  Do not add a `dist/` step — it breaks the git-URL install model.
- **Branch model:** work on `main`. Don't open per-feature branches unless asked.

## Deploy / release

This is a **versioned library** — releases are git tags. There is no service to
deploy; consumers pin to a tag in their `dependencies`:

```jsonc
"@synappsis/email-provider": "github:SynappsisAI/email-provider#v0.1.0"
```

To cut a release: bump `package.json` `version` to match the tag **in the same
commit**, move `## [Unreleased]` changelog entries under `## [X.Y.Z] — <date>`,
commit, then tag `vX.Y.Z` and push the tag. Because consumers pin, a release is not
live until they bump their pin — coordinate a contract change with synmail and
AgentFleet.

## Conventions & gotchas

- **Public API takes/returns `Buffer` for file bytes — never base64.** Any base64
  is an internal wire-format detail of a provider. Don't leak it into `types.ts`.
- **Keep it pure mechanics.** Auth, policy, transport, and rate-limiting-per-caller
  are out of scope by design — push back if a change would pull them in.
- **A contract change (`types.ts` / `EmailProvider`) is a BREAKING release.** Bump
  the major, flag it `BREAKING —` in the changelog, and coordinate with consumers.
- Provider setup prerequisites (Google OAuth scopes, Microsoft app permissions) live
  in `README.md` — update there, not here.


<!-- BEGIN synappsis-shared: changelog (v1) — managed by /synappsis-claude-md; edit the skill asset, not this block -->
## Changelog (`CHANGELOG.md`) — shared Synappsis practice

Every repo keeps a `CHANGELOG.md` at its root. It has two consumers: people working
in the repo, and a job that **once a day aggregates all repos' changelogs into
`synappsis-brain`** — it pulls what's new via
`git diff <last-processed-SHA>..HEAD -- CHANGELOG.md`, so keep the file clean and
never add "already ingested" markers.

**When to update it:** in the **same commit** as the change. If the commit changes
behavior, capabilities, prod data/config, or anything a teammate should know → it
gets an entry. Refactors, typos, formatting and trivial bumps don't.

**Format** — [Keep a Changelog](https://keepachangelog.com), adapted:
- **Newest-first sections, each headed by a release id + ISO date.** Use the mode
  that fits the repo (recorded in this CLAUDE.md's repo-specific section):
  - *Continuously-deployed* repos (long-lived branch, no release tags — most
    services / infra): `## [2026-05-20]` — the date is the release id.
  - *Versioned* repos (apps / libraries that tag SemVer releases):
    `## [1.4.0] — 2026-05-20`, accumulating new entries under `## [Unreleased]`
    between releases.
  Always include the ISO date — the daily aggregation orders cross-repo by it.
- Under each section, category subsections: `Added` · `Changed` · `Fixed` ·
  `Deprecated` · `Removed` · `Security` · `Ops` (infra / data / config / deploy
  changes that aren't code).
- Each entry is **self-contained**, bold-title first: **what** changed + **why** +
  **impact** in 1–3 sentences. The brain ingests it **without the repo's context**,
  so it must stand alone (never "fixed yesterday's bug").
- Prefix an entry with **`BREAKING —`** when it is not backward-compatible (removed/
  renamed field, config-param / payload / contract change, changed default). It is
  the highest-priority signal; in versioned repos it is what forces a major version
  bump, and the daily brain summary surfaces these first.
- Mark **→ prod** when the change is already deployed. Cite the short commit SHA
  (`` `a1b2c3d` ``) only when it points to a *prior* commit; an entry introduced in
  its **own** commit does not cite a SHA (it changes on commit / amend).
- Write entries in **Spanish** (Synappsis changelog standard).

**Cross-repo:** log each change in the repo that **owns** it — a prod data / infra
change goes in the infra repo's changelog, not a module's, so the daily aggregation
does not double-count it. Cross-reference other repos by name.

**Not a `git log`:** it is curated — one entry per meaningful change, grouped by
impact, not 1:1 with commits.
<!-- END synappsis-shared: changelog -->
<!-- BEGIN synappsis-shared: house-rules (v2) — managed by /synappsis-claude-md; edit the skill asset, not this block -->
## How Claude should behave in Synappsis repos

- **Confirm before prod-affecting actions** (deploy, config changes, data changes).
  Most repos have no staging — smoke-test with rollback when you can.
- **Commit and push only when the user asks.** Work on the repo's long-lived branch;
  don't open per-feature branches unless asked. The one exception is `/handoff`, which
  commits and pushes its own handoff document (see the Session handoffs block).
- **Never commit secrets** (`.env*`, `*.pem`, `*.p12`) or real customer / PII data
  (anonymize fixtures).
- **End commit messages** with the `Co-Authored-By` trailer.
- **Update `CHANGELOG.md`** in the same commit as any change worth an entry (see the
  Changelog section).
- **Language:** conversation with the user is in Spanish; code, comments, commit
  messages and docs (including this file) are in English; `CHANGELOG.md` entries are
  in Spanish.
<!-- END synappsis-shared: house-rules -->
<!-- BEGIN synappsis-shared: handoff (v2) — managed by /synappsis-claude-md; edit the skill asset, not this block -->
## Session handoffs (`handoff-history/`)

Handoffs are the continuity record between sessions — written so a fresh agent can
resume the work without asking the user anything. Two skills bookend a session:
`/read-handoff` loads the most recent handoff and resumes where the last one left off;
`/handoff` captures the current state into the next one when work is still in flight.

- Handoffs live in `handoff-history/HANDOFF-<N>.md` at the repo root, numbered
  sequentially from `1`. `/handoff` creates the directory and picks the next `N`;
  `/read-handoff` reads the highest `N` (compared numerically).
- Each one is self-contained: goal, current state, what's done and why, concrete next
  steps, relevant paths / commands / branches / gotchas, and a "Suggested skills"
  section. Reference existing artifacts (PRDs, plans, issues, diffs) by path or URL
  instead of duplicating them; redact secrets and PII.
- `/handoff` commits and pushes the new document immediately — this is the one
  explicit checkpoint exempt from "commit and push only when asked".
- **Handoffs are not `CHANGELOG.md` entries.** The `handoff-history/` directory and
  its git history are their own record; keep them out of the changelog so the daily
  `synappsis-brain` aggregation stays limited to product changes.
<!-- END synappsis-shared: handoff -->
<!-- BEGIN synappsis-shared: brain-access (v1) — managed by /synappsis-claude-md; edit the skill asset, not this block -->
## Synappsis company knowledge (the brain)

When a task needs company context you can't derive from this repo — clients, projects,
products, strategy, pricing, or past decisions — consult the **Synappsis brain**, the
operational knowledge base at `github.com/SynappsisAI/synappsis-brain`.

Work from a canonical local mirror so this is identical on every machine:

- **Ensure the mirror exists and is current** before reading:
  ```bash
  if [ -d ~/.claude/synappsis-brain/.git ]; then
    git -C ~/.claude/synappsis-brain pull --rebase --autostash origin main
  else
    git clone https://github.com/SynappsisAI/synappsis-brain.git ~/.claude/synappsis-brain
  fi
  ```
- **Read** from `~/.claude/synappsis-brain/`. Start at its `CLAUDE.md` (auto-generated
  dashboard) and `AGENTS.md` (canonical orientation), then the synthesized layer:
  `clientes/`, `proyectos/`, `productos/`, `decisiones/`, `estrategia/`.
- **The brain is read-only from here.** Never write to or hand-edit the mirror. Only the
  AgentFleet agents (Jarvis, Lonchito, etc.) write to the brain; capturing work from this
  repo into it is out of scope.

Don't copy brain content into this repo; reference it by path instead.
<!-- END synappsis-shared: brain-access -->

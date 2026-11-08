---
name: green-review-before-commit
description: "Every code change needs a green adversarial review before commit; iterate review→fix rounds, but STOP and give Rafa options when findings oscillate or contradict"
metadata:
  node_type: memory
  type: feedback
  originSessionId: 3705a875-310c-42e8-a356-eb4cc372370e
  modified: 2026-11-04T18:36:04.088Z
---

Every code change must come with a **green adversarial review** (no outstanding valid findings)
before it is treated as done or committed. Multiple review rounds are fine: review → fix →
re-review until clean.

**Exception (Rafa, 2026-11-04):** if the findings start to **oscillate or contradict each other**
across rounds (round N asks to remove X, round M asks to put it back; or two rounds demand
incompatible designs), **stop the loop** — don't fix-and-rerun — and present Rafa the options
(what each side of the contradiction proposes, trade-offs, a recommendation) for him to decide.

**Why:** Rafa asked for green reviews explicitly; tests + typecheck alone are not enough. But on
issue #10 the loop ran 12 rounds, and the later ones mostly flip-flopped (Gmail Message-ID
read-back removed then re-requested; wrap errors vs. keep their class; delete vs. keep a draft
after an ambiguous send), so continuing just churns code.

**How to apply:** after `npm run typecheck` / `npm test` pass, run an adversarial review of the
diff (e.g. the `code-review` skill). Fix confirmed findings and re-review. Track each finding
across rounds; the first time one contradicts or reverses an earlier round, stop and escalate
with options instead of launching another round. Items needing live verification (e.g. real
Graph behavior) are reported as such, not looped on. Report review rounds, fixes, and declined
items before proposing commit/tag/push.

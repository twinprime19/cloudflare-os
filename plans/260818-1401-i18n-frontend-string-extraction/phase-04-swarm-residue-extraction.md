---
phase: 4
title: "Swarm Residue Extraction"
status: completed
effort: "1-2 days (mostly parallel agent time)"
---

# Phase 4: Swarm Residue Extraction

## Overview

Cheapest-model (haiku-class) parallel agents wrap what the codemod couldn't judge, driven by `codemod-residue-report.json`. Semantic work: concatenations → single interpolated messages, manual "item(s)" → ICU plural, toast/error strings, helper-built strings. Orchestrated by main session; strict file ownership per agent.

## Requirements

- Functional: every residue-report entry either wrapped correctly or explicitly escalated (never silently skipped, never guessed).
- Non-functional: cost-minimal (cheapest model, scoped prompts, ~200K agent context respected); no two agents edit the same file (orchestration-protocol rule).

## Architecture

- Sharding: group residue by file, ~10–15 files per agent, balanced by candidate count. `ChatInterface.tsx` (VERIFIED 8,020 lines, 200+ strings, 76 `useKumoToastManager` call sites repo-wide) = ONE dedicated agent working alone on that file (sequential internally; no line-range splitting — same-file parallel edits forbidden). Sequencing DECIDED (validation 2026-08-18): proceed now; pending plan 260815-1442 rebases onto wrapped code if/when it lands.
<!-- Updated: Validation Session 1 - ChatInterface scale + sequencing decision -->
- Conventions doc (written first, fed to every agent): `plans/260818-1401-i18n-frontend-string-extraction/extraction-conventions.md` — wrapping forms with before/after examples: `<Trans>` for JSX, `` t`...` `` for attrs/plain strings, `plural()` / ICU for counts, interpolation for concatenation (`"Deleted " + n + " files"` → one message with placeholder — NEVER sentence fragments), what NOT to touch (identifiers, URLs, error `code` values, backend-originated strings rendered verbatim, console/log, LLM prompt text, test assertions).
- Escalation channel: each agent returns `{wrapped: n, escalated: [{file, line, reason}]}`; ambiguous cases (is this string user-facing? does this concat cross a function boundary?) go to escalation, resolved by high-quality model afterward.
- Verification per shard (cheap + deterministic): build must pass, `detect.ts` count for the shard's files must reach zero-or-allowlisted, `lingui extract` diff reviewed for malformed messages.

## Related Code Files

- Create: `extraction-conventions.md` (this plan dir)
- Modify: residue files across `packages/workshop-frontend/src` (per-shard ownership)
- Modify: `packages/i18n/locales/en.po` (regenerated after each wave)

## Implementation Steps

1. High-quality model writes `extraction-conventions.md` from Phase 3 residue patterns; include 6–10 worked examples covering every rule.
2. Partition residue report into shards; record shard→files ownership table in this file.
3. Launch wave 1 (a few shards) as calibration: review results closely, tighten conventions doc if agents misapply rules; then launch remaining shards in parallel.
4. Per-shard verification (build + detector + extract diff). Failed shard → fix-forward by re-prompting THAT agent with the specific findings (change context per orchestration rules — no blind retries).
5. Dedicated ChatInterface.tsx agent runs after calibration wave, with extra examples relevant to chat UI (toasts, slash commands, markdown-adjacent copy).
6. High-quality model resolves all escalations.
7. Aggregate review: code-reviewer subagent over sampled semantic transforms (concats/plurals — the risky category), not the mechanical bulk.

## Success Criteria

- [x] Detector reports zero unwrapped candidates in workshop-frontend (or entries on the documented allowlist) — guard exits 0 with "No unlocalized user-facing strings found"; zero allowlist entries were added to reach it.
- [x] All escalations resolved; none silently dropped — ledger below.
- [x] Build, typecheck, frontend tests green after every wave; final extract deterministic — 809 messages, byte-identical across consecutive extracts, 0 obsolete entries.
- [x] No sentence-fragment messages in catalog (spot-check concatenation rewrites) — went beyond a spot-check: all 809 msgids were scanned for fragment shapes. 21 flagged, 19 cleared with a recorded reason each, 2 genuine fragments found and fixed (`providers.tsx` label/value split, `WorkpiecePicker.tsx` leading separator). An earlier independent sweep of the then-578-message catalog found zero.

### Escalation ledger (closed by the orchestrator, 2026-08-19)

Every escalation raised by a shard, and where it ended. None dropped.

| Origin | Escalation | Resolution |
|---|---|---|
| ChatInterface `:3264-3265` | pluralized by appending `"s"` to the runtime noun `logKind` | `logKind` proved to be a finite set; one `<Plural>` per value, byte-identical English verified |
| AdminPage `:753` | `\n` escape in a placeholder; template-literal equivalence unsafe to assume | converted after empirical verification; the `\n` is present in the extracted msgid |
| ChatInterface `:7219/:7221` | `actor` = `isOwnChange ? "You" : author.name` would ship "You" untranslated inside a placeholder | per-actor branch duplication + the function-call `plural()` macro (returns `string`, so no `ReactNode` change); verified byte-identical across 2 actor branches × 4 counts |
| OutOfCreditsModal `:104-114`, `:121-134` | marked mandatory-escalate in the conventions doc | resolved in-shard at the sentence boundary; 3+4 whole-sentence messages instead of a 2×2 cross-product of near-duplicates |
| Shards R2, R3, R4, E | — | 0 escalated (R4 additionally allowlisted 1 dead-code entry, broken out from `wrapped` so the arithmetic stays honest) |

Categories deliberately NOT wrapped, each with a written reason in `scripts/i18n-codemod/unlocalized-allowlist.txt`, are exceptions rather than escalations: dead prototype trees, backend `.includes()` match needles, CSS/Tailwind constants, telemetry discriminants, capnweb RPC validation errors surfacing only inside a sandboxed third-party iframe, the `DiffStatus` enum, and one agent-prompt mirror cross-referenced to `overseer.ts:5853`.

## Risk Assessment

- Cheap-model misjudgment on semantics → calibration wave + escalation channel + deterministic verification; the guard catches misses forever after.
- Plural/interpolation rewrites change rendered output → these are the ONLY behavior-touching edits; concentrate reviewer attention here; English ICU plural must reproduce current strings exactly.
- Agent collisions → hard file-ownership table; orchestrator enforces.
- Token blowout → shards scoped to file lists + conventions doc only (no conversation history), cheapest model, per orchestration-protocol context isolation.

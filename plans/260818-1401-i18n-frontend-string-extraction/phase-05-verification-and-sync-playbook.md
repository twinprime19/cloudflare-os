---
phase: 5
title: "Verification and Sync Playbook"
status: completed
effort: "0.5 day"
---

# Phase 5: Verification and Sync Playbook

## Overview

Prove done-ness end to end, wire the recurring protections into the repo's check chain, and document how future upstream syncs and new features keep the extraction healthy. Executed by high-quality model.

## Requirements

- Functional: full green pipeline; `i18n:check` script fails CI-style when catalog is stale or bare strings appear; docs complete.
- Non-functional: translation-readiness proven WITHOUT translating (per approved scope): a throwaway pseudo-locale round-trip verifies the pipeline, then is deleted.

## Architecture

- `i18n:check` (root script, `scripts/check-i18n.ts`) = hash catalogs → `lingui extract` → hash again → compare, plus a locale-configuration check, plus the detector run. Wired into the root `lint` chain between `types:scripts` and `types:check`.

  `git diff --exit-code` was the planned comparison and is **unsound here**, so it is not used. `git diff` reports on tracked paths only, and `packages/i18n/` is untracked on this branch, so that command exits 0 against a catalog that has drifted arbitrarily far — a check that looks like coverage and provides none. Verified directly: appending a line to `packages/i18n/locales/en.po` and running `git diff --exit-code -- packages/i18n/locales/en.po` exits 0. Hashing does not depend on git, so it holds on an untracked tree, on a fresh clone and in CI alike. Git is still consulted for two things hashing cannot see: an **ignored** catalog fails the check (it can never be committed, so no clone or CI run would ever have it), and an **untracked** one prints a note explaining that the hash comparison — not git — is what verified the result.
- `docs/i18n.md`: how to write new strings (macro forms + ICU examples), how to run extract/check, locale-switching convention for the future (root re-render `key={locale}`), Kumo known gap, backend adoption path (out of scope, pointer to brainstorm report), and the upstream-sync playbook.
- Sync playbook (the recurring cost, made mechanical): on conflict at a wrapped line → take upstream's copy → re-wrap (run codemod on conflicted files — it is idempotent and handles the mechanical cases; residue via detector) → `i18n:extract` → changed English becomes new untranslated entry (correct behavior by design).

## Related Code Files

- Create: `docs/i18n.md` (owned by a concurrent docs agent, not by this phase)
- Create: `scripts/check-i18n.ts` + `scripts/check-i18n.test.ts` (colocated, `node --test`)
- Modify: root `package.json` (`i18n:check` script; into the `lint` chain)
- Modify: `packages/i18n/src/index.tsx` — comments only. The `loadCatalog` docstring claimed adding `locales/vi.po` "is the whole change", which the pseudo-locale round-trip disproved; the recipe now lives on `SUPPORTED_LOCALES` along with the measured silent-failure mode.
- Modify: `plans/reports/brainstorm-260818-1401-i18n-frontend-string-extraction-report.md` (mark next-steps progress)

## Implementation Steps

1. Full pipeline: `pnpm exec vp run -r --no-cache build`, `pnpm test` (scripts tests incl. codemod), frontend `vitest run`, `pnpm lint`, `pnpm i18n:check`.

   Uncached deliberately, and spelled with `vp` rather than `pnpm run build --no-cache`: the latter forwards the flag to `tsc`, which dies with TS5023. It matters because `vp` replays a cached `tsc` result over a change set that added files, so a cached green build says nothing about the files this work introduced.
2. Pseudo-locale round-trip: generate temporary `xx.po` (machine-mangled English, e.g. bracketed `[!! ... !!]`), activate locally, user visually confirms strings flip in dev server → proves zero-source-change translation path → delete `xx.po`.
3. Write `docs/i18n.md` (≤800 lines per docs cap; will be far under).
4. Sweep: grep plan files + code for stale terms/contradictions (whole-plan consistency gate); resolve.
5. User visual smoke of the full app in English; user decides on posting the upstream Discussion draft (in brainstorm report).
6. STOP — ask user to verify before any commit (standing rule: no commits without explicit request).

## Success Criteria

- [ ] All acceptance criteria in plan.md check off — 4 of 7 ticked. The three left open are the human visual smoke, `docs/i18n.md` (another agent's file), and the "zero source changes" wording, which the measurement contradicts. See plan.md for each.
- [x] Pseudo-locale flip demonstrated and cleaned up — an unmodified component rendered `[!! Reconnecting… !!]`; interpolation survived the flip. `locales/pseudo.po`, the throwaway test and the temporary `lingui.config.ts` line are all removed, and `git status`'s untracked set differs from the pre-phase snapshot only by `scripts/check-i18n.ts` and `scripts/check-i18n.test.ts`.
- [x] `i18n:check` in lint chain; deliberately-staled catalog fails it — proven four ways: a catalog with one message block removed → exit 1; the catalog deleted outright → exit 1 ("absent before it"); a planted bare string → exit 1; a catalog for an unconfigured locale → exit 1. Clean tree → exit 0.
- [ ] `docs/i18n.md` written with sync playbook — concurrent docs agent; not verified by this phase.
- [ ] User verification requested; no commit made without explicit instruction — nothing was committed, staged or pushed. Handoff for user verification is the orchestrator's step.

## Risk Assessment

- `i18n:check` slows lint chain → measured. Warm chain: 4.78–5.33s without, 6.34–6.48s with, so roughly +1.4s (~+28% of a warm run, and noise next to a cold build). `pnpm i18n:check` alone is 1.47–1.69s, of which ~0.5s is `lingui extract` and ~0.35s the detector; the rest is pnpm process startup.
- Pseudo-locale accidentally committed → did not happen; `git status` compared against a snapshot taken before the phase started.
- Docs drift → `docs/i18n.md` is the single i18n doc; future phases (backend adoption) get their own brainstorm/plan.

## Outcome (2026-08-19)

`lingui extract` preserves the `POT-Creation-Date` header of an existing catalog, so a clean
re-extract is byte-identical and the hash comparison does not flap. It writes a fresh header only
when it creates the file, which is already a failure state for this check. It also does **not**
delete a catalog for a locale it was never configured with, which is why the locale-configuration
check exists rather than relying on `--clean`.

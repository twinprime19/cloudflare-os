---
phase: 1
title: "Lingui Integration Spike"
status: completed
effort: "0.5 day"
---

# Phase 1: Lingui Integration Spike

## Overview

Prove Lingui v6 (target 6.6.0, current stable, published 2026-07-24 — passes the 24h release-age gate; all five packages verified on registry) works in this repo's exact toolchain before any mass extraction: `@vitejs/plugin-react` v5 (Babel) + vite 7.3.6 (pinned, esbuild transform) + vite-plus runner + TypeScript 7 (tsgo) native typecheck + vitest 4. Wrap 2–3 real strings, build, test, extract. Executed by high-quality model.
<!-- Updated: Validation Session 1 - Lingui v5 → v6.6.0; vite config verified -->

### Decision: **GO** — macros (2026-08-18)

Lingui v6.6.0 macros compile cleanly under this exact chain. No fallback to the non-macro runtime API (`msg`/`defineMessage` + `i18n._`) is needed; Phase 3's codemod targets **macro** output.

Evidence (worktree `feat/i18n-core`, spike component `src/components/AutoApproveConfirmDialog.tsx`, 3 strings):

- `pnpm build` (tsgo typecheck + vite build, whole workspace) → exit 0. tsgo resolves `@lingui/core/macro` and `@lingui/react/macro` under `moduleResolution: bundler` with no errors and no shims.
- `pnpm -C packages/workshop-frontend run test:run` → exit 0, 28 files / 150 tests, no new failures.
- `pnpm run lint:check` (oxlint) → exit 0, 0 errors.
- `lingui extract` → exactly 3 messages in `src/locales/en.po`, interpolation emitted as ICU `{actionLabel}`.
- Production bundle: 0 residual `@lingui/*/macro` imports across 26 JS chunks; compiled shape is `jsx(Trans, { id: "p-9AxK", message: "Always approve “{actionLabel}”?", values: { actionLabel } })`.
- Babel pass costs ~0.66s on a 5.53s frontend build (+12%); plugin-react's Babel step was previously skipped entirely on build.

Two wiring facts Phase 2 must carry over (both cost a build failure or a blank UI if dropped):

1. `lingui()` needs an explicit `cwd` — `vp run -r` resolves vite configs from the workspace root, and Lingui's config discovery is `process.cwd()`-based, so the default fails the whole task graph with "No Lingui config found".
2. `descriptorFields: 'message'` on the Babel macro plugin. The v6 default `'auto'` strips a production build to `id` only, so every wrapped string renders as a bare hash unless a matching compiled catalog is loaded. Keeping the message makes inline English its own fallback.

Full detail: `plans/reports/from-fullstack-developer-to-orchestrator-260818-1519-i18n-phase01-lingui-spike-report.md`

## Requirements

- Functional: `<Trans>` and `` t`...` `` compile, render English, and `lingui extract` emits them into an `en` catalog.
- Non-functional: zero behavior change; new deps pass `minimumReleaseAge: 1440` (Lingui 6.6.0 published 2026-07-24 — satisfied).

## Architecture

- Macro transform rides `@vitejs/plugin-react`'s `babel.plugins` option (`@lingui/babel-plugin-lingui-macro`) — scoped to workshop-frontend only; does NOT touch the workspace-wide "transform stays on esbuild" constraint (that concerns decorator lowering in worker packages; plugin-react's Babel pass is already active for this package).
- `@lingui/vite-plugin` for importing compiled `.po` catalogs.
- tsgo typecheck: macros import from `@lingui/react/macro` / `@lingui/core/macro` which ship real type defs; typecheck never runs Babel, so types must pass standalone.

## Related Code Files

- Modify: `packages/workshop-frontend/package.json` (deps: `@lingui/core`, `@lingui/react`; dev: `@lingui/cli`, `@lingui/babel-plugin-lingui-macro`, `@lingui/vite-plugin`)
- Modify: `packages/workshop-frontend/vite.config.ts` (VERIFIED — plugins: `TanStackRouterVite`, `react()`, `tailwindcss()`, `tsconfigPaths()`; add lingui babel plugin into `react()` options + `@lingui/vite-plugin`)
- Create: `packages/workshop-frontend/lingui.config.ts` (temporary spike location; moves to `packages/i18n` in Phase 2)
- Modify: ONE low-traffic component (pick something small, e.g. an error/empty-state view) — wrap 2–3 strings
- Modify: app entry (provider not required for spike if using global `i18n.activate('en')` — smallest possible footprint)

## Implementation Steps

1. Create worktree `feat/i18n-core` off `dev`; copy `.dev.vars*` from main checkout.
2. Add dependencies (`pnpm add` in workshop-frontend). Verify lockfile passes release-age policy.
3. `lingui.config.ts`: `locales: ['en']`, `sourceLocale: 'en'`, format `po`, catalog path.
4. Wire babel plugin into plugin-react config; add `@lingui/vite-plugin`.
5. Bootstrap minimal i18n activation (`i18n.load('en', messages)`; `i18n.activate('en')`) at entry.
6. Wrap 2–3 strings: one JSX text via `<Trans>`, one attribute via `` t`...` `` (core macro, global instance), one with interpolation.
7. Run: `pnpm --filter @gadgets/workshop-frontend exec lingui extract` → catalog contains the 3 messages with correct placeholders.
8. Run `pnpm build` (includes tsgo typecheck) + `pnpm --filter @gadgets/workshop-frontend test:run`.
9. User verifies rendering via dev server (user-run only).

## Success Criteria

- [ ] Build + typecheck + frontend tests green with macros in place
- [ ] `lingui extract` emits exactly the wrapped messages; English renders unchanged
- [ ] Vite dev + prod build both transform macros (verify prod bundle has no macro imports)
- [ ] GO/NO-GO decision recorded in this file's Overview

## Risk Assessment

- Macro transform fails under this Vite/plugin chain → FALLBACK (decided in brainstorm): Lingui non-macro runtime API (`msg`/`defineMessage` + `i18n._`), same architecture, slightly larger diffs. Decision MUST be recorded here before Phase 3, because codemod output shape depends on it.
- tsgo (TS 7.0.2) type incompatibility with Lingui types → same fallback path; report exact errors.
- vite-plus wraps/owns vite config in unexpected ways → read `vite-plus` docs/config in repo, adjust wiring; worst case pin plugin config at root `vite.config.ts`.

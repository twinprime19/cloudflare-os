---
phase: 2
title: "i18n Core Package and Wiring"
status: completed
effort: "1 day"
---

# Phase 2: i18n Core Package and Wiring

## Overview

Promote the spike into a proper additive core: `packages/i18n` owns config, catalogs, and the configured i18n instance; workshop-frontend consumes it. Add the unlocalized-string guard. Executed by high-quality model. Depends on Phase 1 GO decision.

## Requirements

- Functional: single source of i18n truth importable by any package later (backend adoption path stays open); locale hardcoded `en` with browser-detection stub; `I18nProvider` at app root.
- Non-functional: package is additive-only (upstream never conflicts inside it); components keep importing macros from Lingui directly (upstream-convergent shape, per brainstorm).

## Architecture

- `packages/i18n` (`@gadgets/i18n`, matching workspace naming): exports configured `i18n` instance, `AppI18nProvider` (wraps `@lingui/react` I18nProvider + loads `en`), `detectLocale()` stub returning `'en'`.
- Catalogs: `packages/i18n/locales/en.po` (single catalog; extraction scans `packages/workshop-frontend/src`). Future `vi.po` lands beside it, zero source changes.
- `lingui.config.ts` lives in `packages/i18n`; root script `i18n:extract` runs the CLI against it.
- Guard (DECIDED, validation 2026-08-18): `scripts/check-unlocalized-strings.ts` — an AST detector (shares detection logic with Phase 3 codemod: same rules for "user-facing string candidate"). DRY: one detection module, two consumers (codemod wraps, checker reports). ESLint path is dead: repo lints with oxlint via Vite+ (rules in root `vite.config.ts`), and `@lingui/eslint-plugin` doesn't exist for v6 anyway.
<!-- Updated: Validation Session 1 - guard = detector script; provider insertion point verified -->

## Related Code Files

- Create: `packages/i18n/package.json`, `packages/i18n/src/index.ts`, `packages/i18n/lingui.config.ts`, `packages/i18n/locales/en.po`
- Modify: `packages/workshop-frontend/package.json` (dep `@gadgets/i18n: workspace:*`); `src/main.tsx` — insert `AppI18nProvider` directly under `<StrictMode>`, ABOVE `FrontendErrorBoundary` (line ~193; decided so error-fallback copy stays translatable later); vite config (keep Phase 1 wiring, point at new config)
- Modify: root `package.json` (script `i18n:extract`, later `i18n:check` in Phase 5)
- Create (conditional): `scripts/check-unlocalized-strings.ts` (repo runs scripts as native TS)
- Delete: spike-local `lingui.config.ts` from workshop-frontend

## Implementation Steps

1. Scaffold `packages/i18n`; move config + catalog; wire exports.
2. Replace spike bootstrap with `AppI18nProvider` at app root; verify Phase 1 wrapped strings still render + extract.
3. Build detector script `scripts/check-unlocalized-strings.ts` (JSX text nodes + whitelisted attrs + sentence-like literals heuristic; allowlist file for intentional exceptions). Investigation done during validation: `vp lint` = oxlint, no ESLint anywhere.
4. Inventory Kumo-internal English strings (grep `node_modules/@cloudflare/kumo` dist for rendered literals like confirm/cancel/toast chrome): document findings as known gap in `docs/i18n.md` draft notes — NOT in scope to fix.
5. Run build + tests + extract; user visual smoke.

## Success Criteria

- [x] `packages/i18n` additive package builds; provider wired; app behavior unchanged
- [x] `pnpm i18n:extract` regenerates `en.po` deterministically
- [x] Guard mechanism chosen, implemented, and flags a deliberately-planted bare string (then remove the plant)
- [x] Kumo string gap documented
- [x] Build/tests green

<!-- Completed 2026-08-18. Report:
plans/reports/from-fullstack-developer-to-orchestrator-260818-1538-i18n-phase02-core-package-report.md
Deltas from the spec above: `lingui.config.ts` stays in `packages/i18n` (as planned), but the Babel
macro plugin reloads Lingui config from `process.cwd()` on every file, so the frontend's vite config
resolves the config once and passes it as `linguiConfig` -- +1 devDep `@lingui/conf`, +1 devDep
`@lingui/format-po` (the `origins: false` option needs the formatter package). The provider really
loads the compiled catalog (`import('../locales/{locale}.po')`) rather than activating an empty one.
Detector heuristics are conservative per Phase 3's spec: a lone word in JSX text is not a candidate,
which Phase 3 owns tuning. Kumo gap documented qualitatively (node_modules scanning blocked);
Phase 5 enumerates it empirically.

Post-review round (2026-08-18, after independent code review): the source-locale catalog is now
STATICALLY imported and activated at module scope. Reason -- `@lingui/core` registers its runtime
ICU compiler only when NODE_ENV !== 'production', so the inline `message` that
`descriptorFields: 'message'` embeds is raw uncompiled ICU: in a production build any interpolated
message resolved from that fallback rendered its placeholder literally. Also fixed:
`lingui extract --clean` (obsolete `#~` entries accumulated otherwise), and catalog `exclude`
patterns anchored at `<rootDir>` (a bare `**/*.test.tsx` matched nothing, so test files were being
extracted into the catalog). -->


## Risk Assessment

- Catalog-per-package pressure later (backend adoption) → `.po` supports multiple catalogs in one config; no rework needed, note in docs.
- Provider re-render semantics with global-instance `t` from core macro: fine while locale is fixed `en`; future locale switching forces re-render via root `key={locale}` — record this convention in docs/i18n.md (Phase 5).

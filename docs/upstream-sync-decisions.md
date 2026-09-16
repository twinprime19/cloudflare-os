# Upstream sync: standing fork decisions

This fork (Contoso `dev`) tracks `cloudflare/cloudflare-os` `main`, which is read-only for us: we sync
from it, we never push, PR, or offer patches to it. Every sync re-collides the same fork features
with the same kind of upstream change, so the reconciliation decisions below are standing policy.
A new sync plan starts by reading this file; it re-opens a decision only with new evidence, and
records the outcome here when it does.

History: syncs on 2026-08-18 (`plans/260818-1030-upstream-main-sync`), 2026-08-19
(`plans/260819-1620-upstream-sync-17-commits`), 2026-08-25
(`plans/260825-1432-upstream-sync-21-commits`), 2026-09-04
(`plans/260904-1548-upstream-sync-50-commits`), 2026-09-09
(`plans/260909-1312-upstream-sync-10-commits`). The i18n-specific mechanics live in
[`i18n.md` → Upstream-sync playbook](./i18n.md#upstream-sync-playbook).

## 1. Shape of a sync

| Rule | Why |
|---|---|
| One `git merge --no-ff` of `upstream/main` on a worktree sync branch cut from `origin/dev`; `dev` is never touched until the user fast-forwards it | dev stays deployable; the branch is disposable |
| Absorb **every** upstream commit, including large rewrites (#275 git storage, #341 step buffer, #420/#421 composer extraction) | cherry-picking around a rewrite makes every later sync worse; later upstream commits sit on top of it |
| Probe with `git merge-tree --write-tree --merge-base=<base> dev upstream/main` before planning; the conflict list and `$TREE` line ranges drive the per-file recipes | conflicts are pre-classified as `i18n-wrap-only`, `logic`, `dep/config`, `insertion-adjacency` |
| Resolve **file by file**; never a blanket `--ours` / `--theirs` | four of thirteen 260904 conflicts were insertion-adjacency where either blanket choice silently deletes one side's feature |
| Audit every fork commit's hunks in the merged result (`git show <sha> -- <file>` → located or MISSING), not only the conflicted regions | a clean auto-merge is not a correct merge (`overseer.ts` +2.5k upstream lines around fork hunks) |
| Independent review before the merge commit; commit on the sync branch only on `COMMIT_OK` **and** explicit user commit authorization; ff / push / cleanup / deploy re-pin each on separate say-so | dev moves only with the user's hand on it |
| `origin/dev` moves during a sync every time (260819, 260825, 260904); re-fetch, merge `origin/dev` into the sync branch, re-check the new commits against the recipes, re-run the quick gate before the ff | recipe drift is the failure mode |

## 2. Who wins

### Upstream wins — take upstream's structure, re-apply the fork's hunks into it

| Surface | Decision | Sync |
|---|---|---|
| Architecture rewrites (storage, agent turn loop, component extraction) | Adopt upstream's shape. Re-home each fork hunk into the new file it belongs to; never keep the old structure alongside | 260825, 260904 |
| English source text at a wrapped line | Take upstream's wording, re-wrap with the codemod, re-extract; a changed sentence becomes a new untranslated msgid by design | 260819 → |
| Identifier renames (`pendingActionsCount` → `pendingActionCount`) | Upstream's name everywhere, including inside a fork conflict hunk | 260904 |
| Dropped imports / derived helpers (`ResourceDescription` moved, `hasProposedChanges` → `chatHasProposedChanges(meta)`) | Upstream's form; the fork's wraps ride on top | 260904 |
| `pnpm-lock.yaml` | Never hand-merged. Take upstream's, make **all** dependency edits, one `pnpm install`, prove with `pnpm install --frozen-lockfile` | 260818 → |
| Root scripts and build tooling (`vp run` forms, `@gadgets/scripts`, `scripts` as a workspace package, `scripts/tsconfig.json` includes) | Upstream's, plus the fork's `i18n:extract` / `i18n:check`; root `lint` keeps `i18n:check` before `types:check` | 260904 |
| Toolchain patterns for **fork-only packages** (`gatekeeper-microsoft`, `i18n`) | Mirror the closest upstream sibling in the same sync (`gatekeeper-mcp` for the configurator import + `@gadgets/scripts` devDep). Upstream's routing guards scan `packages/*` and will fail on an unmigrated fork package | 260904 |
| Dependency versions | `catalog:` pins in fork-only packages; never a literal version that can drift from the catalog | 260904 |
| Workers `compatibility_date` / toolchain bumps (#448: every package `2026-02-02` → `2026-09-04`, `wrangler` catalog, regenerated `worker-configuration.d.ts`) | Mirror into `gatekeeper-microsoft` in the same sync: bump `wrangler.jsonc` (`wrangler.dev.jsonc` is a gitignored dev-server artefact — nothing to edit), regenerate types with **`pnpm types:generate`** (`scripts/generate-worker-types.ts`; raw `wrangler types` points `mainModule` at `.wrangler/validate/…` and breaks `tsc`), and bump the fork's `gatekeeper-microsoft` entry in `scripts/release/testdata/golden-manifest.json` — `manifest-lib.test.ts` diffs generated vs golden and fails on a date mismatch. The auto-merged lockfile keeps the fork-only importer on the old `wrangler`; the one `pnpm install` repoints it | 260909 |
| `scripts/run-dev-server.ts` | The fork's shape stays (`ADMINS` from root `.dev.vars`, Microsoft gatekeeper seeding); upstream's diff is ported into `dev-server-config.ts` + tests when it touches config, accepted as-is when it is lifecycle only | 260825, 260904 |

### Fork wins — the fork's decision survives; upstream's version of the same thing is dropped

| Surface | Decision | Guard |
|---|---|---|
| Suggested Cloudflare model list (`SUGGESTED_MODEL_CATALOG.cloudflare` in `workshop-shared/src/api.ts`) | The fork **curates** it: exactly the models the user chose, in the user's order (first = default). Upstream model additions are dropped on sync unless the user adds them. When both sides add the same model, keep the fork's entry and comment, delete upstream's duplicate key. Never lower a published `contextWindow`. Upstream packages that derive a model-id type from the catalog can also hard-code a dropped id (260904: `workshop-evals/src/config.ts` + its tests named `deepseek-v4-pro-0813` / `glm-5.2`); re-point them at the fork's carried models, keep default ≠ override so the override test still bites | `workshop-backend/__tests__/ai-models-unified-catalog.test.ts` pins the exact key list and windows; `pnpm build` (`EvalModelId` derives from the catalog, so a dropped id is a type error) |
| `COMPACTION_TRIGGER_RATIO = 0.65` | Plain ratio, single knob; no cap or derived constants | `agent-compaction.test.ts` |
| Chat attachments: 8 files / 1.75 MiB each / 8 MiB per message; eight document types convert to markdown (≤10 MiB): pdf (provider-conditional), docx, xlsx, xls, xlsm, xlsb, odt, ods — text only, `images: {convert: false}` always (PR #24, 2026-09-09); `.doc`/`.pptx`/`.numbers` rejected; per-message document budget; convertible documents upload **one at a time** (backpressure on the Workers AI conversion in one Durable Object) | Constants live in `chat-attachment-validation.ts`, `overseer.ts`, and the frontend composer (`attachments/prepareChatAttachment.ts`, `useComposerAttachments.ts`); when upstream moves the composer, the caps and the conversion gate move with it into the new files. 260909: upstream's `ComposerAddMenu` took over the upload button; its `onUpload` clicks the same file input, so the gate needs no re-home — verify the wiring, don't re-add the old button | backend `chat-attachment-*.test.ts`; frontend: upstream's own composer tests adapted to the fork's constants |
| Admin shared models, per-model availability toggles (hard-off), provider endpoint rename without re-entering the token | Fork feature set stays; upstream UI rewrites get the fork's `disabled={!admin}` and toggles re-applied. Upstream test harnesses that register a model hit PR #9's admin gate too: `addModel` throws for a non-admin caller, and `setPreferredModel` is dead (client-side model persistence retired, every turn names its model). 260904: `integration-tests/src/agent-session.ts`'s `openAgentSession` registers through a new `offerSharedModel()` calling `AdminApi.addSharedModel()` as the harness's admin, and drops the `setPreferredModel` setup call | `admin-config.test.ts`, `admin-shared-models.test.ts`; `pnpm test` (harness typechecks and the integration suite) |
| Vietnamese locale, `detectLocale`, Lingui wraps across `workshop-frontend/src` | Fork-carried on every sync; vi fully translated **before** dev fast-forwards | `pnpm i18n:check`; `po-catalog.test.ts` count |
| User-message copy button | Keep both: the fork's on the user-message hover footer and upstream's on assistant code blocks — complementary | — |
| Gatekeeper login-profile hints (Microsoft Graph avatar/name seed) | Additive contract in `workshop-shared/src/gatekeeper.ts`; keep-both on insertion-adjacency conflicts | `login-profile.test.ts` |
| Microsoft 365 gatekeeper (mail, Outlook attachment import, Teams, SharePoint List) | Fork-only package; upstream never touches it, but its build/test wiring must follow upstream's current pattern (see above) | its own suite + the deploy build invocation |
| Share-link placeholder regex (`getPlaceholderRanges` in `workshop-frontend/src/resourceMatching.ts`) | The fork's named-parameter pattern ends `(?![\w:])`, so a SharePoint/OneDrive "Copy link" segment (`/:l:/`, `/:w:/`, `/:x:/`) reads as part of the URL rather than as a `:placeholder` to fill in. The file is upstream-owned and this is the fork's only hunk in it: re-apply the trailing-colon exclusion whenever a sync rewrites that regex | `resourceMatching.test.ts` |
| Gadget source-zip download (`fflate`, `downloadSourceZip`) | Re-homed onto whatever file resolver upstream's editor uses so archive and editor never disagree | `GadgetCodeInterface.test.tsx`, `fileTransfers.test.ts` |
| AI Gateway `user.id` metadata (`ai-models.ts`), 408-no-body error capture (`describeRequestFailure`), Home "New format" shortcuts | Fork hunks; auto-merge but are grep-verified every sync | Phase 3 grep table in each sync plan |
| CI: Bonk workflows | Dormant in the fork (`workflow_dispatch`-only, triggers commented out in upstream's own #258 shape). The fork has no `ask-bonk` App and no `CF_AI_GATEWAY_*` secrets; a live Bonk is red on every PR | grep `on:` blocks |
| Deploy contract | `~/repos/contoso-os-deploy` is hands-off. Its `deploy.mjs` invocations (`vp run -F <pkg> build` per deployed package, `FORMAT_BLUEPRINTS_DIR`) are run verbatim as gates in the sync worktree; findings and a staging checklist are reported, never applied | the exact `deploy.mjs` commands |

## 3. Tests during a sync

- **Adapt, don't add.** Existing tests are adapted to upstream's new shape or the fork's constants. No new test files, harnesses, or invariant tests are introduced during a sync unless a fork feature would otherwise have no guard at all (user rule, 2026-09-04: "whatever in the fork is working, don't add new things unless absolutely necessary").
- Never weaken, skip, or delete a test to get green. A test that can no longer express its invariant under the new architecture is reported, not loosened.
- Upstream tests that mount a React subtree fail until their root comes from `createI18nRoot()` (`workshop-frontend/src/testI18nRoot.tsx`); swap the shared harness once rather than each consumer.
- Packages without a test script (`workshop-shared`, `i18n`) are not failures; do not add empty suites.
- The fork-hunk tests added on 2026-08-25 (`packages/workshop-backend/__tests__/*` pinning overseer caps, admin-model resolution, profile seeding, the built-in model routing — qwen at the time, GLM 5.3 Flash since PR #20) exist precisely so every later sync is self-checking; keep them green, extend them only when a fork hunk moves.
- `MAX_COMPOSER_ATTACHMENTS` and the 8 MiB per-message total have no biting frontend assertion, because upstream ships no frontend test of its own `5`/lower caps to adapt; the backend `chat-attachment-*.test.ts` suite is the only guard. Do not add a frontend assertion during a sync — that is a new test, not an adaptation.

## 4. Recurring traps (check every sync)

| Trap | What it looks like | Check |
|---|---|---|
| Same addition on both sides | Both add the same object key (`glm-5.3-flash`, 260904); git auto-merges both → duplicate key, `tsc` error, exact-list test fails | grep the key count after merge |
| Rename inside a conflict hunk | The fork hunk still uses the old name that upstream renamed everywhere else | `pnpm build` plus a grep for the old name |
| Insertion adjacency | Both sides inserted next to the same comment/field; either blanket resolution deletes a feature | keep-both, then grep both sides' symbols |
| Fork-only package vs a workspace-wide upstream guard | `build-gatekeeper-configurator.test.ts` routing guard, `env-passthrough.test.ts` `EXPECTED` map, `release/manifest-lib.ts` short-name rules all scan `packages/*` | run the scripts tests; migrate the fork package the way the sibling did |
| Vite+ task cache replays a stale pass | `pnpm test` green after an edit that should fail | iterate with `test:run`, use cached `pnpm test` only for the final proof |
| Fresh worktree missing `typed-storage/dist` | its `exports` resolve to `dist/index.js` | `pnpm build` before any test |
| Local Durable Object state | a storage migration (#275) is one-way per workspace | smoke on a **copy** of `.wrangler/state`; never symlink or `--persist-to` the main checkout |
| Format blueprints | #265 unpacked `.gadget` archives into `format-blueprints/<name>/blueprint.json` + `files/`; legacy pairs still build; migrate with `pnpm import:format-blueprint <export.gadget> --new <name>` | downstream blueprint branches (DNTT, Contoso Slides) |
| `po-catalog.test.ts` entry count | pins the vi catalog size | update to the real post-extract number, never delete |
| Upstream hard-codes a curated-away model id | `tsc` error on a model-id union, or an exact-list test fails, in an upstream package like `workshop-evals` | `pnpm build` plus a grep for the dropped id |
| Widened RPC contract vs a fork-only gatekeeper test | capnweb `@validateRpc` rejects a fork test's old call arity after upstream adds a parameter (260904: `applyAction(action, cache: RpcStub<GitCache>)` broke `gatekeeper-microsoft/__tests__`) | run the package suite; fix on the test side with the same throwing stub upstream's own `gatekeeper-cloudflare` tests use, never in the implementation |

## 5. Execution model

- **opus**: every code edit, conflict resolution, re-home, test adaptation, translation, commit on the sync branch, and the independent merge review.
- **sonnet**: read-only — merge-tree scouting, gate execution and reporting, deploy-contract audit, translation consistency reads.
- Subagents own disjoint files; one agent owns config/lockfile and the single `pnpm install`; `scripts/i18n-codemod/unlocalized-allowlist.txt` is controller-only (agents report entries with a reason).
- Gate set named in every brief: `pnpm install --frozen-lockfile`, `pnpm build`, `pnpm test`, per-package `test:run`, `pnpm exec vp lint`, `pnpm types:scripts`, `pnpm i18n:check`, and the `deploy.mjs` build invocations. `@gadgets/integration-tests` is user-run.
- Handoff to the user leads with the worktree path, branch, head sha, and the `pnpm dev-server --port <port>` line; each worktree gets its own port and a copy of `.dev.vars*`.

## 6. Reopening a decision

Reopen only with new evidence (a measured failure, a changed upstream contract, a product change the user asked for). Record the change in this file with the date and the sync plan that made it. The suggested-model list, attachment caps, and compaction ratio are product settings: a sync never changes them on its own; the user changes them in a normal feature PR (as PR #20 did on 2026-09-04), and the next sync protects the new value.

// Gate for the frontend's i18n extraction. Three independent failures, each of which silently rots
// a localized app, and none of which any other check in this repo would notice:
//
//   1. The catalogs under `packages/i18n/locales` no longer match what `lingui extract` produces
//      from the source. A translator's `vi.po` is keyed off the source catalog, so a stale `en.po`
//      means new copy has no entry to translate and removed copy keeps one forever.
//   2. A catalog exists for a locale that `lingui.config.ts` does not list. Measured: the app then
//      loads that catalog, reports the locale as active, and renders every string in English --
//      see `unconfiguredCatalogs`.
//   3. A user-facing string is in the source with no i18n macro around it. Delegated whole to
//      `check-unlocalized-strings.ts`, which owns every rule about what counts.
//
//   node scripts/check-i18n.ts
//
// Exit 1 when the check fails, 2 when it could not run.
//
// Freshness is decided by hashing the catalogs, running `lingui extract`, and hashing again -- NOT
// by `git diff --exit-code` after the extract, which is the obvious implementation and is unsound
// here. `git diff` reports on tracked paths only, so while `packages/i18n/` is untracked (it is new
// on this branch) that command exits 0 against a catalog that has drifted arbitrarily far: a check
// that looks like coverage and provides none. Hashing does not care whether git has ever heard of
// the file, so it holds on an untracked working tree, on a fresh clone, and in CI alike.
//
// Two conditions git is still consulted for, because hashing cannot see them, and both are about
// the catalog's future rather than its current contents -- see `trackingVerdict`.
//
// The extract runs in place, so a stale catalog is left corrected on disk. That is deliberate: the
// failure message then asks for a review and a commit rather than for another command.
import { createHash } from "node:crypto";
import { existsSync, readFileSync, readdirSync } from "node:fs";
import { basename, join, relative, resolve } from "node:path";
import { spawnSync } from "node:child_process";
import { pnpmCommand } from "./pnpm-command.ts";

/** The repository root. This file lives in `scripts/`. */
const REPO_ROOT = join(import.meta.dirname, "..");

/** The package that owns `lingui.config.ts` and every catalog. */
const I18N_PACKAGE = join(REPO_ROOT, "packages", "i18n");

/** Directory names never descended into when looking for catalogs. */
const SKIPPED_DIRECTORIES = new Set(["node_modules", ".git", "dist"]);

/** `pnpm` arguments that regenerate every catalog. Mirrors the root `i18n:extract` script. */
const EXTRACT_ARGS = ["--filter", "@gadgets/i18n", "run", "extract"];

/** The Lingui configuration, which decides which locales the build compiles catalogs for. */
const LINGUI_CONFIG = join(I18N_PACKAGE, "lingui.config.ts");

/**
 * `locales: [...]` and `sourceLocale: '...'` in `lingui.config.ts`, read as text.
 *
 * Text rather than an import: the config imports `@lingui/cli`, which resolves from the i18n
 * package and not from here, and evaluating it would mean building a module graph to read two array
 * literals. Both declarations are single-line literals, so the parse is exact for the shapes that
 * exist -- and it returns `undefined` rather than a guess for anything else, which
 * {@link main} treats as a failure. A config this cannot read is a config nothing here can verify.
 */
export function configuredLocales(configSource: string): string[] | undefined {
  const locales = /^\s*locales:\s*\[([^\]]*)\]/m.exec(configSource);
  if (!locales) return undefined;
  const quoted = [...locales[1].matchAll(/['"]([^'"]+)['"]/g)].map((match) => match[1]);
  if (quoted.length === 0) return undefined;
  const source = /^\s*sourceLocale:\s*['"]([^'"]+)['"]/m.exec(configSource);
  return source ? [...new Set([source[1], ...quoted])] : quoted;
}

/**
 * The catalogs whose locale `lingui.config.ts` does not list, given repo-relative `.po` paths.
 *
 * This is the check for a measured silent failure, not a hypothetical one. `@lingui/vite-plugin`
 * compiles a `.po` whose locale is absent from `locales` by falling back to the source language:
 * the import succeeds, the module exports the full set of message ids, `i18n.activate()` reports
 * the new locale as active, and every string on screen is still English. Nothing throws and nothing
 * logs. The only symptom is a translation that does not appear, which reads as a bad catalog rather
 * than as a missing line of config.
 */
export function unconfiguredCatalogs(
  catalogPaths: string[],
  configured: readonly string[],
): string[] {
  const known = new Set(configured);
  return catalogPaths.filter((path) => !known.has(basename(path, ".po"))).toSorted();
}

/** What git knows about a catalog file. */
export type CatalogTracking = "tracked" | "untracked" | "ignored" | "unknown";

/** How the catalogs on disk changed across an extraction. Paths are relative to the repo root. */
export interface CatalogDrift {
  /** Catalogs the extraction created. */
  added: string[];
  /** Catalogs the extraction deleted. */
  removed: string[];
  /** Catalogs whose contents the extraction rewrote. */
  changed: string[];
}

/** A problem worth reporting, and whether it should fail the run. */
export interface Verdict {
  fatal: boolean;
  message: string;
}

/**
 * Every `.po` file under `directory`, as repo-relative path to SHA-256 of its bytes.
 *
 * Recursive rather than a listing of `locales/`, so that repointing `lingui.config.ts` at a new
 * catalog directory inside the package is seen as an addition instead of leaving this comparing a
 * file nothing writes any more. A config that moves catalogs out of the package entirely is beyond
 * this: `packages/i18n/src/index.tsx` imports `../locales/en.po` by path, so that change breaks the
 * build, which is a louder failure than anything here.
 *
 * An absent directory yields an empty map rather than throwing, because "no catalog before the
 * extract" is a result this check reports on rather than an error -- see {@link main}.
 */
export function snapshotCatalogs(directory: string): Map<string, string> {
  const catalogs = new Map<string, string>();
  if (!existsSync(directory)) return catalogs;

  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) {
      if (SKIPPED_DIRECTORIES.has(entry.name)) continue;
      for (const [nested, hash] of snapshotCatalogs(path)) catalogs.set(nested, hash);
    } else if (entry.isFile() && entry.name.endsWith(".po")) {
      catalogs.set(
        relative(REPO_ROOT, path),
        createHash("sha256").update(readFileSync(path)).digest("hex"),
      );
    }
  }

  return catalogs;
}

/** What changed between two {@link snapshotCatalogs} results. */
export function compareCatalogs(
  before: Map<string, string>,
  after: Map<string, string>,
): CatalogDrift {
  const added: string[] = [];
  const changed: string[] = [];
  for (const [path, hash] of after) {
    const previous = before.get(path);
    if (previous === undefined) added.push(path);
    else if (previous !== hash) changed.push(path);
  }
  const removed = [...before.keys()].filter((path) => !after.has(path));
  return { added: added.toSorted(), removed: removed.toSorted(), changed: changed.toSorted() };
}

/** Whether a {@link CatalogDrift} reports anything at all. */
export function hasDrift(drift: CatalogDrift): boolean {
  return drift.added.length + drift.removed.length + drift.changed.length > 0;
}

/**
 * The failure text for a drifted catalog.
 *
 * `added` is spelled out separately from `changed` because it usually means something other than
 * "a locale was added": on a fresh clone or in CI it means the catalog is not committed, so the
 * extract produced it from nothing and there was never anything to compare against.
 */
export function formatDrift(drift: CatalogDrift): string {
  const lines: string[] = ["The message catalogs are not in sync with the source."];
  for (const path of drift.changed) lines.push(`  changed by the extract: ${path}`);
  for (const path of drift.added) {
    lines.push(`  produced by the extract, absent before it: ${path}`);
  }
  for (const path of drift.removed) lines.push(`  deleted by the extract: ${path}`);
  if (drift.added.length > 0) {
    lines.push(
      "",
      "A catalog that did not exist until the extract ran cannot have been verified against " +
      "anything. If this is a clone or CI, the catalog is missing from the commit.",
    );
  }
  lines.push(
    "",
    "The extract ran in place, so the corrected catalogs are already on disk. Review the diff " +
    "and commit it.",
  );
  return lines.join("\n");
}

/**
 * What to say about git's view of the catalogs, or `undefined` when it has nothing to add.
 *
 * Hashing decides freshness on its own, so none of this gates correctness of the comparison that
 * just ran. It gates whether that comparison will still mean anything tomorrow:
 *
 * - `ignored` is fatal. An ignored catalog can never be committed, so every clone and every CI run
 *   starts without it, no translator can ever receive it, and the only reason a check like this
 *   would go green is that it regenerated the file it was about to inspect.
 * - `untracked` is a warning. It is the state of a catalog that is written but not yet committed --
 *   true of this whole package right now -- and it resolves itself on the next commit. It is
 *   reported rather than ignored because it is exactly the state in which a `git diff`-based
 *   freshness check would pass while verifying nothing.
 * - `unknown` is a warning: no git, or no repository. Nothing is wrong, there is just nothing to say.
 */
export function trackingVerdict(states: Map<string, CatalogTracking>): Verdict | undefined {
  const withState = (state: CatalogTracking): string[] =>
    [...states].filter(([, value]) => value === state).map(([path]) => path);

  const ignored = withState("ignored");
  if (ignored.length > 0) {
    return {
      fatal: true,
      message:
        `git ignores these catalogs, so they can never be committed:\n` +
        ignored.map((path) => `  ${path}`).join("\n") +
        "\n\nEvery clone and every CI run would start without them, and this check would only be " +
        "inspecting files it had just regenerated. Remove the ignore rule.",
    };
  }

  const untracked = withState("untracked");
  if (untracked.length > 0) {
    return {
      fatal: false,
      message:
        `Note: git does not track these catalogs yet:\n` +
        untracked.map((path) => `  ${path}`).join("\n") +
        "\n\nFreshness above was decided by hashing, which does not depend on git, so the result " +
        "holds. A `git diff --exit-code` check would have passed here regardless of the contents. " +
        "Commit the catalogs so a clone has them.",
    };
  }

  if (withState("unknown").length > 0) {
    return { fatal: false, message: "Note: could not ask git about the catalogs (no repository?)." };
  }

  return undefined;
}

/** Runs `git` with `args` from the repo root, or `undefined` when it cannot be run. */
function git(args: string[]): string | undefined {
  const result = spawnSync("git", args, { cwd: REPO_ROOT, encoding: "utf8" });
  // `check-ignore` exits 1 to mean "nothing matched", which is an answer rather than a failure;
  // `ls-files` exits 0 either way. Anything above 1, or a spawn error, means git did not answer.
  if (result.error || result.status === null || result.status > 1) return undefined;
  return result.stdout;
}

/** The non-empty lines of a git command's output, as a set. */
function pathSet(output: string): Set<string> {
  return new Set(output.split("\n").map((line) => line.trim()).filter(Boolean));
}

/** What git knows about each of `paths`, which are relative to the repo root. */
function trackingStates(paths: string[]): Map<string, CatalogTracking> {
  const tracked = git(["ls-files", "--", ...paths]);
  const ignored = git(["check-ignore", "--", ...paths]);
  const states = new Map<string, CatalogTracking>();
  if (tracked === undefined || ignored === undefined) {
    for (const path of paths) states.set(path, "unknown");
    return states;
  }

  const trackedPaths = pathSet(tracked);
  const ignoredPaths = pathSet(ignored);
  for (const path of paths) {
    if (trackedPaths.has(path)) states.set(path, "tracked");
    else if (ignoredPaths.has(path)) states.set(path, "ignored");
    else states.set(path, "untracked");
  }
  return states;
}

/** Runs a child process attached to this one's streams, and reports whether it succeeded. */
function run(label: string, command: string, args: string[]): boolean {
  const result = spawnSync(command, args, { cwd: REPO_ROOT, stdio: "inherit" });
  if (result.error) {
    console.error(`Could not run ${label}: ${result.error.message}`);
    return false;
  }
  if (result.status !== 0) {
    console.error(`\n${label} failed (exit ${result.status ?? "signal " + result.signal}).`);
    return false;
  }
  return true;
}

function main(): void {
  if (process.argv.length > 2) {
    console.error(`Unexpected argument: ${process.argv[2]}. This check takes none.`);
    process.exitCode = 2;
    return;
  }

  const before = snapshotCatalogs(I18N_PACKAGE);

  const [command, args] = pnpmCommand(EXTRACT_ARGS);
  if (!run("lingui extract", command, args)) {
    process.exitCode = 2;
    return;
  }

  const after = snapshotCatalogs(I18N_PACKAGE);
  if (after.size === 0) {
    console.error(
      `\nThe extract produced no catalogs under ${relative(REPO_ROOT, I18N_PACKAGE)}. ` +
      "There is nothing to verify, so this check cannot pass.",
    );
    process.exitCode = 1;
    return;
  }

  let failed = false;

  const drift = compareCatalogs(before, after);
  if (hasDrift(drift)) {
    console.error(`\n${formatDrift(drift)}`);
    failed = true;
  } else {
    console.log(`\n${after.size} catalog(s) already match the source.`);
  }

  const locales = configuredLocales(readFileSync(LINGUI_CONFIG, "utf8"));
  if (locales === undefined) {
    console.error(
      `\nCould not read the locale list out of ${relative(REPO_ROOT, LINGUI_CONFIG)}. ` +
      "Whether every catalog will be compiled with its own translations therefore cannot be " +
      "verified, so this check cannot pass.",
    );
    failed = true;
  } else {
    const unconfigured = unconfiguredCatalogs([...after.keys()], locales);
    if (unconfigured.length > 0) {
      console.error(
        `\n${relative(REPO_ROOT, LINGUI_CONFIG)} does not list a locale for these catalogs:\n` +
        unconfigured.map((path) => `  ${path}`).join("\n") +
        `\n\nThe build compiles them from the source language instead of from their own ` +
        "translations: the app would activate the locale and still render English, with nothing " +
        `thrown and nothing logged. Add the locale to \`locales\` (currently ` +
        `${JSON.stringify(locales)}), and to \`SUPPORTED_LOCALES\` in ` +
        "packages/i18n/src/index.tsx so it can be selected.",
      );
      failed = true;
    }
  }

  const verdict = trackingVerdict(trackingStates([...after.keys()]));
  if (verdict) {
    if (verdict.fatal) failed = true;
    (verdict.fatal ? console.error : console.log)(`\n${verdict.message}`);
  }

  if (!run("the unlocalized-string check", process.execPath, [
    join(REPO_ROOT, "scripts", "check-unlocalized-strings.ts"),
  ])) {
    failed = true;
  }

  if (failed) process.exitCode = 1;
}

if (process.argv[1] && import.meta.filename === resolve(process.argv[1])) main();

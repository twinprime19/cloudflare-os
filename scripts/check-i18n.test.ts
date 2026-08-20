import assert from "node:assert/strict";
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { basename, join } from "node:path";
import { tmpdir } from "node:os";
import { after, describe, it } from "node:test";
import {
  compareCatalogs, configuredLocales, formatDrift, hasDrift, snapshotCatalogs, trackingVerdict,
  unconfiguredCatalogs, type CatalogTracking,
} from "./check-i18n.ts";

/** Temporary trees the snapshot tests build, removed together at the end. */
const temporaryRoots: string[] = [];

after(() => {
  for (const root of temporaryRoots) rmSync(root, { recursive: true, force: true });
});

/** A directory tree from `relative path -> contents`, with the directories created as needed. */
function tree(files: Record<string, string>): string {
  const root = mkdtempSync(join(tmpdir(), "check-i18n-"));
  temporaryRoots.push(root);
  for (const [path, contents] of Object.entries(files)) {
    const target = join(root, path);
    mkdirSync(join(target, ".."), { recursive: true });
    writeFileSync(target, contents);
  }
  return root;
}

/** The snapshot keyed by file name, so assertions do not depend on the temporary root's path. */
function byName(snapshot: Map<string, string>): Map<string, string> {
  return new Map([...snapshot].map(([path, hash]) => [basename(path), hash]));
}

describe("snapshotCatalogs", () => {
  it("hashes every .po file and nothing else", () => {
    const snapshot = byName(snapshotCatalogs(tree({
      "locales/en.po": "msgid \"Cancel\"\n",
      "locales/vi.po": "msgid \"Cancel\"\n",
      "lingui.config.ts": "export default {}\n",
      "locales/en.js": "export const messages = {}\n",
    })));
    assert.deepEqual([...snapshot.keys()].toSorted(), ["en.po", "vi.po"]);
  });

  it("gives identical bytes the same hash and different bytes a different one", () => {
    const first = byName(snapshotCatalogs(tree({ "locales/en.po": "same\n" })));
    const same = byName(snapshotCatalogs(tree({ "locales/en.po": "same\n" })));
    const other = byName(snapshotCatalogs(tree({ "locales/en.po": "different\n" })));
    assert.equal(first.get("en.po"), same.get("en.po"));
    assert.notEqual(first.get("en.po"), other.get("en.po"));
  });

  // A dependency ships hundreds of `.po` files. Hashing them would make every install look like
  // catalog drift, and the walk itself would dominate the check's runtime.
  it("does not descend into dependency or build directories", () => {
    const snapshot = byName(snapshotCatalogs(tree({
      "locales/en.po": "ours\n",
      "node_modules/some-package/locales/en.po": "theirs\n",
      "dist/en.po": "built\n",
    })));
    assert.deepEqual([...snapshot.keys()], ["en.po"]);
  });

  // "No catalogs at all" is a result this check reports on, not a crash: it is what a clone looks
  // like when the catalogs were never committed.
  it("returns nothing for a directory that does not exist", () => {
    assert.equal(snapshotCatalogs(join(tmpdir(), "check-i18n-absent-directory")).size, 0);
  });
});

describe("compareCatalogs", () => {
  const before = new Map([["a.po", "1"], ["b.po", "2"]]);

  it("reports no drift when every hash survives the extract", () => {
    const drift = compareCatalogs(before, new Map(before));
    assert.deepEqual(drift, { added: [], removed: [], changed: [] });
    assert.equal(hasDrift(drift), false);
  });

  it("reports a rewritten catalog as changed", () => {
    const drift = compareCatalogs(before, new Map([["a.po", "1"], ["b.po", "CHANGED"]]));
    assert.deepEqual(drift.changed, ["b.po"]);
    assert.equal(hasDrift(drift), true);
  });

  it("reports a catalog the extract had to create", () => {
    const drift = compareCatalogs(new Map(), new Map([["a.po", "1"]]));
    assert.deepEqual(drift.added, ["a.po"]);
    assert.equal(hasDrift(drift), true);
  });

  it("reports a catalog the extract removed", () => {
    const drift = compareCatalogs(before, new Map([["a.po", "1"]]));
    assert.deepEqual(drift.removed, ["b.po"]);
    assert.equal(hasDrift(drift), true);
  });
});

describe("formatDrift", () => {
  it("says a created catalog was never verified against anything", () => {
    const message = formatDrift({ added: ["locales/en.po"], removed: [], changed: [] });
    assert.match(message, /absent before it: locales\/en\.po/);
    assert.match(message, /cannot have been verified against anything/);
  });

  it("points a stale catalog at the corrected file rather than at another command", () => {
    const message = formatDrift({ added: [], removed: [], changed: ["locales/en.po"] });
    assert.match(message, /changed by the extract: locales\/en\.po/);
    assert.match(message, /already on disk/);
    assert.doesNotMatch(message, /never verified/);
  });
});

describe("configuredLocales", () => {
  it("reads the locale list, and folds in the source locale", () => {
    assert.deepEqual(
      configuredLocales("export default defineConfig({\n  sourceLocale: 'en',\n" +
        "  locales: ['en', 'vi'],\n})\n"),
      ["en", "vi"],
    );
  });

  it("reads a list the source locale is not already in", () => {
    assert.deepEqual(
      configuredLocales("  sourceLocale: 'en',\n  locales: [\"vi\"],\n"),
      ["en", "vi"],
    );
  });

  it("reads a list with no sourceLocale beside it", () => {
    assert.deepEqual(configuredLocales("  locales: ['en'],\n"), ["en"]);
  });

  // A shape this cannot read must not read as "no locales configured", which would fail every
  // catalog. `main` turns `undefined` into a failure of its own that names the config file.
  it("returns undefined rather than guessing at a shape it cannot read", () => {
    assert.equal(configuredLocales("export default defineConfig({})\n"), undefined);
    assert.equal(configuredLocales("  locales: [],\n"), undefined);
    assert.equal(configuredLocales("  locales: SUPPORTED,\n"), undefined);
  });

  // The list Lingui is actually configured with today, so the guard is exercised against the real
  // file's shape and not only against fixtures.
  it("reads the repository's own config", () => {
    const config = readFileSync(
      join(import.meta.dirname, "..", "packages", "i18n", "lingui.config.ts"), "utf8");
    assert.deepEqual(configuredLocales(config), ["en", "vi"]);
  });
});

describe("unconfiguredCatalogs", () => {
  it("accepts a catalog whose locale is configured", () => {
    assert.deepEqual(unconfiguredCatalogs(["packages/i18n/locales/en.po"], ["en"]), []);
  });

  // Measured behaviour: the app loads such a catalog, activates the locale, and renders English.
  it("reports a catalog with no locale entry behind it", () => {
    assert.deepEqual(
      unconfiguredCatalogs(["packages/i18n/locales/vi.po", "packages/i18n/locales/en.po"], ["en"]),
      ["packages/i18n/locales/vi.po"],
    );
  });
});

/** A `path -> state` map, spelled as an object literal. */
const states = (entries: Record<string, CatalogTracking>): Map<string, CatalogTracking> =>
  new Map(Object.entries(entries));

describe("trackingVerdict", () => {
  it("stays quiet when git tracks every catalog", () => {
    assert.equal(trackingVerdict(states({ "locales/en.po": "tracked" })), undefined);
  });

  // The condition that makes a catalog permanently unverifiable everywhere but this one machine.
  it("fails on an ignored catalog", () => {
    const verdict = trackingVerdict(states({ "locales/en.po": "ignored" }));
    assert.equal(verdict?.fatal, true);
    assert.match(verdict.message, /locales\/en\.po/);
    assert.match(verdict.message, /Remove the ignore rule/);
  });

  // Untracked is transient -- the state of a catalog written but not yet committed -- so it must
  // not fail the run. It is still reported, because it is precisely the state in which the obvious
  // `git diff --exit-code` implementation of this check passes while verifying nothing.
  it("warns without failing on an untracked catalog, and says why it still matters", () => {
    const verdict = trackingVerdict(states({ "locales/en.po": "untracked" }));
    assert.equal(verdict?.fatal, false);
    assert.match(verdict.message, /git diff --exit-code/);
    assert.match(verdict.message, /decided by hashing/);
  });

  it("prefers the ignored catalog's message when both states are present", () => {
    const verdict = trackingVerdict(states({ "a.po": "untracked", "b.po": "ignored" }));
    assert.equal(verdict?.fatal, true);
    assert.match(verdict.message, /b\.po/);
  });

  it("notes without failing when git could not answer", () => {
    const verdict = trackingVerdict(states({ "locales/en.po": "unknown" }));
    assert.equal(verdict?.fatal, false);
    assert.match(verdict.message, /could not ask git/);
  });
});

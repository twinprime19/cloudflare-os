import assert from "node:assert/strict";
import { mkdirSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, it } from "node:test";
import {
  parseAllowlist,
  validateCoverage,
  validateFragment,
  validateTranslation,
  type Fragment,
} from "./merge.ts";
import { parsePo, serializePo } from "./po-catalog.ts";

const NO_ALLOWLIST = new Map<string, string>();

function rules(findings: { rule: string }[]): string[] {
  return findings.map((finding) => finding.rule);
}

describe("validateTranslation", () => {
  it("accepts a faithful Vietnamese translation with collapsed plural", () => {
    const msgid =
      "{0, plural, one {# other person loses} other {# other people lose}} access through {1}. Keep anyone?";
    const msgstr =
      "{0, plural, other {# người khác mất}} quyền truy cập qua {1}. Giữ lại ai không?";
    assert.deepEqual(validateTranslation(msgid, msgstr, NO_ALLOWLIST), []);
  });

  it("rejects an empty msgstr", () => {
    assert.deepEqual(rules(validateTranslation("Cancel", "  ", NO_ALLOWLIST)), ["empty"]);
  });

  it("rejects a dropped placeholder", () => {
    const findings = validateTranslation("Delete {name}?", "Xóa chứ?", NO_ALLOWLIST);
    assert.deepEqual(rules(findings), ["placeholder-parity"]);
    assert.match(findings[0].detail, /missing name/);
  });

  it("rejects an invented placeholder", () => {
    const findings = validateTranslation("Delete?", "Xóa {name}?", NO_ALLOWLIST);
    assert.deepEqual(rules(findings), ["placeholder-parity"]);
    assert.match(findings[0].detail, /unexpected name/);
  });

  it("rejects a plural that drops the octothorpe everywhere", () => {
    const findings = validateTranslation(
      "{0, plural, one {# file} other {# files}}",
      "{0, plural, other {các tệp}}",
      NO_ALLOWLIST,
    );
    assert.deepEqual(rules(findings), ["placeholder-parity"]);
    assert.match(findings[0].detail, /missing #0/);
  });

  it("rejects dropped or unbalanced tags", () => {
    assert.deepEqual(
      rules(validateTranslation("<0>Learn more</0>", "Tìm hiểu thêm", NO_ALLOWLIST)),
      ["tag-parity"],
    );
    assert.deepEqual(
      rules(validateTranslation("<0>Learn more</0>", "<0>Tìm hiểu thêm", NO_ALLOWLIST)),
      ["tag-parity"],
    );
    assert.deepEqual(
      rules(validateTranslation("<0>a</0> <1>b</1>", "<0>a</0> <0>b</0>", NO_ALLOWLIST)),
      ["tag-parity"],
    );
  });

  it("accepts reordered tags that stay balanced", () => {
    assert.deepEqual(
      validateTranslation("<0>save</0> or <1>discard</1>", "<1>bỏ</1> hoặc <0>lưu</0>", NO_ALLOWLIST),
      [],
    );
  });

  it("rejects syntactically invalid ICU", () => {
    const findings = validateTranslation("{0} files", "{0, plural, other {tệp}", NO_ALLOWLIST);
    assert.deepEqual(rules(findings), ["icu-invalid"]);
  });

  it("rejects a translated plural that keeps the English `one` branch", () => {
    const findings = validateTranslation(
      "{0, plural, one {# file} other {# files}}",
      "{0, plural, one {# tệp} other {# tệp}}",
      NO_ALLOWLIST,
    );
    assert.deepEqual(rules(findings), ["vi-plural"]);
    assert.match(findings[0].detail, /only "other" and "=N"/);
  });

  it("accepts exact-number branches beside other", () => {
    assert.deepEqual(
      validateTranslation(
        "{0, plural, one {# file} other {# files}}",
        "{0, plural, =1 {Một tệp} other {# tệp}}",
        NO_ALLOWLIST,
      ),
      [],
    );
  });

  it("rejects an untranslated English leftover unless allowlisted", () => {
    assert.deepEqual(rules(validateTranslation("Webhook", "Webhook", NO_ALLOWLIST)), [
      "english-leftover",
    ]);
    const allowlist = new Map([["Webhook", "technical term kept in English"]]);
    assert.deepEqual(validateTranslation("Webhook", "Webhook", allowlist), []);
  });

  it("rejects msgstr that is not NFC-normalized", () => {
    // Force the decomposed form (base letters plus combining marks).
    const decomposed = "d\u1EF1 \u00E1n".normalize("NFD");
    assert.notEqual(decomposed, decomposed.normalize("NFC"));
    const findings = validateTranslation("project", decomposed, NO_ALLOWLIST);
    assert.deepEqual(rules(findings), ["nfc"]);
  });
});

describe("validateCoverage", () => {
  const shard = ["a", "b", "c"];

  it("accepts translated plus escalated covering the shard exactly", () => {
    const fragment: Fragment = {
      translations: { a: "x", b: "y" },
      escalated: [{ msgid: "c", reason: "ambiguous referent" }],
    };
    assert.deepEqual(validateCoverage(shard, fragment), []);
  });

  it("rejects a missing msgid", () => {
    const findings = validateCoverage(shard, { translations: { a: "x", b: "y" }, escalated: [] });
    assert.deepEqual(rules(findings), ["coverage"]);
    assert.equal(findings[0].msgid, "c");
  });

  it("rejects an msgid outside the shard", () => {
    const fragment: Fragment = {
      translations: { a: "x", b: "y", c: "z", d: "invented" },
      escalated: [],
    };
    const findings = validateCoverage(shard, fragment);
    assert.deepEqual(rules(findings), ["coverage"]);
    assert.equal(findings[0].msgid, "d");
  });

  it("rejects overlap between translated and escalated", () => {
    const fragment: Fragment = {
      translations: { a: "x", b: "y", c: "z" },
      escalated: [{ msgid: "c", reason: "unsure" }],
    };
    assert.deepEqual(rules(validateCoverage(shard, fragment)), ["coverage"]);
  });

  it("rejects an escalation without a reason", () => {
    const fragment: Fragment = {
      translations: { a: "x", b: "y" },
      escalated: [{ msgid: "c", reason: " " }],
    };
    assert.deepEqual(rules(validateCoverage(shard, fragment)), ["coverage"]);
  });
});

describe("validateFragment", () => {
  it("combines coverage and per-translation findings", () => {
    const findings = validateFragment(
      ["Delete {name}?", "Cancel"],
      { translations: { "Delete {name}?": "Xóa?" }, escalated: [] },
      NO_ALLOWLIST,
    );
    assert.deepEqual(rules(findings).toSorted(), ["coverage", "placeholder-parity"]);
  });
});

describe("parseAllowlist", () => {
  it("parses tab-separated entries and skips comments", () => {
    const allowlist = parseAllowlist("# comment\nAPI\tuniversal technical term\n\nOK\tuniversal\n");
    assert.equal(allowlist.get("API"), "universal technical term");
    assert.equal(allowlist.size, 2);
  });

  it("rejects a line without a reason", () => {
    assert.throws(() => parseAllowlist("API\n"), /no tab-separated reason/);
  });
});

describe("merge --check and --apply through the CLI", () => {
  const MERGE = join(import.meta.dirname, "merge.ts");

  function setUpPlanDir(): { dir: string; catalogPath: string } {
    const dir = mkdtempSync(join(tmpdir(), "i18n-merge-test-"));
    mkdirSync(join(dir, "shards"));
    mkdirSync(join(dir, "waves"));
    const catalogSource =
      'msgid ""\n"Language: vi\\n"\nmsgstr ""\n\n' +
      '#. placeholder {name}: user.name\nmsgid "Delete {name}?"\nmsgstr ""\n\n' +
      'msgid "Cancel"\nmsgstr ""\n';
    const catalogPath = join(dir, "vi.po");
    writeFileSync(catalogPath, catalogSource);
    writeFileSync(
      join(dir, "shards", "shard-01.json"),
      JSON.stringify({
        shard: "01",
        messages: [
          { msgid: "Delete {name}?", hints: ["placeholder {name}: user.name"] },
          { msgid: "Cancel", hints: [] },
        ],
      }),
    );
    return { dir, catalogPath };
  }

  async function runMerge(args: string[]): Promise<{ code: number; output: string }> {
    const { execFile } = await import("node:child_process");
    return new Promise((resolvePromise) => {
      execFile("node", [MERGE, ...args], (error, stdout, stderr) => {
        const code = typeof error?.code === "number" ? error.code : error ? 1 : 0;
        resolvePromise({ code, output: stdout + stderr });
      });
    });
  }

  it("checks a fragment and merges it, leaving escalations empty until resolved", async () => {
    const { dir, catalogPath } = setUpPlanDir();
    const fragmentPath = join(dir, "waves", "shard-01.vi.json");
    writeFileSync(
      fragmentPath,
      JSON.stringify({
        translations: { "Delete {name}?": "Xóa {name}?" },
        escalated: [{ msgid: "Cancel", reason: "testing the escalation path" }],
      }),
    );
    const check = await runMerge(["--check", fragmentPath]);
    assert.equal(check.code, 0, check.output);

    const applyBefore = await runMerge([
      "--apply", "--shards", join(dir, "shards"), "--waves", join(dir, "waves"), "--catalog", catalogPath,
    ]);
    assert.equal(applyBefore.code, 0, applyBefore.output);
    assert.match(applyBefore.output, /1 entries remain untranslated/);

    writeFileSync(
      join(dir, "waves", "escalations.vi.json"),
      JSON.stringify({ translations: { Cancel: "Hủy" }, escalated: [] }),
    );
    const applyAfter = await runMerge([
      "--apply", "--shards", join(dir, "shards"), "--waves", join(dir, "waves"), "--catalog", catalogPath,
    ]);
    assert.equal(applyAfter.code, 0, applyAfter.output);

    const catalog = parsePo(readFileSync(catalogPath, "utf8"));
    assert.deepEqual(
      catalog.entries.map((entry) => entry.msgstr),
      ["Xóa {name}?", "Hủy"],
    );
    // The merged catalog still parses and re-serializes stably.
    assert.equal(serializePo(catalog), readFileSync(catalogPath, "utf8"));

    const audit = await runMerge(["--audit", "--catalog", catalogPath]);
    assert.equal(audit.code, 0, audit.output);
  });

  it("refuses to merge a fragment with findings and reports them", async () => {
    const { dir, catalogPath } = setUpPlanDir();
    writeFileSync(
      join(dir, "waves", "shard-01.vi.json"),
      JSON.stringify({
        translations: { "Delete {name}?": "Xóa?", Cancel: "Hủy" },
        escalated: [],
      }),
    );
    const result = await runMerge([
      "--apply", "--shards", join(dir, "shards"), "--waves", join(dir, "waves"), "--catalog", catalogPath,
    ]);
    assert.equal(result.code, 1);
    assert.match(result.output, /placeholder-parity/);
    // Nothing was written: the catalog still has empty msgstr.
    const catalog = parsePo(readFileSync(catalogPath, "utf8"));
    assert.deepEqual(catalog.entries.map((entry) => entry.msgstr), ["", ""]);
  });

  it("audits English leftovers against the allowlist", async () => {
    const { dir, catalogPath } = setUpPlanDir();
    writeFileSync(
      join(dir, "waves", "shard-01.vi.json"),
      JSON.stringify({
        translations: { "Delete {name}?": "Xóa {name}?", Cancel: "Cancel" },
        escalated: [],
      }),
    );
    const allowlistPath = join(dir, "allowlist.txt");
    writeFileSync(allowlistPath, "Cancel\tpretend brand term for this test\n");
    const apply = await runMerge([
      "--apply", "--shards", join(dir, "shards"), "--waves", join(dir, "waves"),
      "--catalog", catalogPath, "--allowlist", allowlistPath,
    ]);
    assert.equal(apply.code, 0, apply.output);

    const auditWithAllowlist = await runMerge([
      "--audit", "--catalog", catalogPath, "--allowlist", allowlistPath,
    ]);
    assert.equal(auditWithAllowlist.code, 0, auditWithAllowlist.output);

    const auditWithout = await runMerge(["--audit", "--catalog", catalogPath]);
    assert.equal(auditWithout.code, 1);
    assert.match(auditWithout.output, /english-leftover/);
  });
});

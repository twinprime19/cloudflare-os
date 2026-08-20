// Validates translation-swarm fragments and merges them into the Vietnamese catalog. Only this
// script (orchestrator-run) ever writes `vi.po`; agents write fragment JSON. Everything checkable
// deterministically is checked here, so model output never reaches the catalog unvalidated.
//
//   node scripts/i18n-translate/merge.ts --check <waves/shard-NN.vi.json> [--shards <dir>]
//   node scripts/i18n-translate/merge.ts --apply --shards <dir> --waves <dir> [--catalog <vi.po>]
//   node scripts/i18n-translate/merge.ts --audit [--catalog <vi.po>]
//
// `--check` validates one fragment against its shard without writing. `--apply` validates every
// fragment and fills the catalog (plus an optional `escalations.vi.json` patch fragment holding
// the orchestrator's resolutions -- validated the same way, minus shard coverage). `--audit`
// reports untranslated and English-identical entries in the catalog; it is also the tripwire for
// upstream syncs that add source strings.
import { existsSync, readFileSync, readdirSync } from "node:fs";
import { createRequire } from "node:module";
import { basename, dirname, join } from "node:path";
import { pathToFileURL } from "node:url";
import { parseArgs } from "node:util";
import { readPoFile, setMsgstr, writePoFile, type PoCatalog } from "./po-catalog.ts";

const REPO_ROOT = join(import.meta.dirname, "..", "..");
const DEFAULT_CATALOG = join(REPO_ROOT, "packages", "i18n", "locales", "vi.po");
const DEFAULT_ALLOWLIST = join(import.meta.dirname, "identical-ok-allowlist.txt");

// --- ICU parsing -------------------------------------------------------------------------------

// `@messageformat/parser` is what lingui itself validates messages with, but it is a transitive
// dependency (packages/i18n -> @lingui/cli -> @lingui/message-utils -> @messageformat/parser), so
// pnpm's strict layout hides it from this script's own resolution. Resolving hop by hop through
// the packages that actually declare each dependency reaches the exact copy lingui uses without
// adding a dependency of our own. `@lingui/message-utils` exports no ".", hence the subpath hop.
function resolveIcuParserPath(): string {
  let location = join(REPO_ROOT, "packages", "i18n", "package.json");
  for (const specifier of [
    "@lingui/cli",
    "@lingui/message-utils/compileMessage",
    "@messageformat/parser",
  ]) {
    location = createRequire(location).resolve(specifier);
  }
  return location;
}

/** One `key {tokens}` branch of a plural/select. */
interface IcuCase {
  key: string;
  tokens: IcuToken[];
}

/** The subset of `@messageformat/parser` token shapes these validators read. */
type IcuToken =
  | { type: "content"; value: string }
  | { type: "argument"; arg: string }
  | { type: "function"; arg: string; key: string }
  | { type: "octothorpe" }
  | { type: "plural" | "select" | "selectordinal"; arg: string; cases: IcuCase[] };

const icuModule = (await import(pathToFileURL(resolveIcuParserPath()).href)) as {
  parse: (source: string) => IcuToken[];
};

/** Parses one ICU message, throwing the parser's own error on invalid syntax. */
export function parseIcu(message: string): IcuToken[] {
  return icuModule.parse(message);
}

// --- Validators --------------------------------------------------------------------------------

/** One validation failure. `rule` names the validator so re-prompts can quote it. */
export interface Finding {
  rule:
    | "fragment-shape"
    | "coverage"
    | "empty"
    | "placeholder-parity"
    | "tag-parity"
    | "icu-invalid"
    | "vi-plural"
    | "english-leftover"
    | "nfc";
  msgid: string;
  detail: string;
}

/** What a swarm agent writes for one shard. */
export interface Fragment {
  translations: Record<string, string>;
  escalated: { msgid: string; reason: string }[];
}

/**
 * Collects each placeholder's name and role. A set (not a multiset) by design: collapsing
 * `one/other` into the single `other` branch Vietnamese uses legitimately changes how many times
 * a nested placeholder occurs, so counting occurrences would reject correct translations. What
 * must survive translation is the set of placeholders and what each one is -- including, per
 * plural argument, whether any branch renders `#`.
 */
function placeholderSignature(
  tokens: IcuToken[],
  into = new Map<string, Set<string>>(),
  enclosingPluralArg?: string,
): Map<string, Set<string>> {
  const add = (name: string, kind: string): void => {
    const kinds = into.get(name) ?? new Set();
    kinds.add(kind);
    into.set(name, kinds);
  };
  for (const token of tokens) {
    if (token.type === "argument") add(token.arg, "argument");
    else if (token.type === "function") add(token.arg, `function:${token.key}`);
    else if (token.type === "octothorpe") {
      // `#` outside a plural is literal text; inside, it renders the nearest plural's number.
      if (enclosingPluralArg !== undefined) add(`#${enclosingPluralArg}`, "octothorpe");
    } else if (token.type === "plural" || token.type === "select" || token.type === "selectordinal") {
      add(token.arg, token.type);
      const pluralArg = token.type === "select" ? enclosingPluralArg : token.arg;
      for (const branch of token.cases) placeholderSignature(branch.tokens, into, pluralArg);
    }
  }
  return into;
}

function comparePlaceholders(msgid: string, msgstr: string): string | undefined {
  const expected = placeholderSignature(parseIcu(msgid));
  const actual = placeholderSignature(parseIcu(msgstr));
  const problems: string[] = [];
  for (const [name, kinds] of expected) {
    const got = actual.get(name);
    if (!got) problems.push(`missing ${name}`);
    else if ([...kinds].toSorted().join(",") !== [...got].toSorted().join(",")) {
      problems.push(`${name} changed role (${[...kinds].join("/")} -> ${[...got].join("/")})`);
    }
  }
  for (const name of actual.keys()) {
    if (!expected.has(name)) problems.push(`unexpected ${name}`);
  }
  return problems.length > 0 ? problems.join("; ") : undefined;
}

/** Lingui's rich-element markers: `<0>`, `</0>`, `<0/>`. Anything else with `<` is plain text. */
const TAG_PATTERN = /<(\/?)(\d+)(\/?)>/g;

function tagSignature(value: string): { counts: Map<string, number>; problem?: string } {
  const counts = new Map<string, number>();
  const stack: string[] = [];
  for (const match of value.matchAll(TAG_PATTERN)) {
    const [, closing, index, selfClosing] = match;
    if (closing && selfClosing) return { counts, problem: `malformed tag ${match[0]}` };
    const key = selfClosing ? `self:${index}` : `open:${index}`;
    if (closing) {
      if (stack.pop() !== index) return { counts, problem: `unbalanced </${index}>` };
    } else {
      if (!selfClosing) stack.push(index);
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }
  }
  if (stack.length > 0) return { counts, problem: `unclosed <${stack[stack.length - 1]}>` };
  return { counts };
}

function compareTags(msgid: string, msgstr: string): string | undefined {
  const expected = tagSignature(msgid);
  const actual = tagSignature(msgstr);
  if (actual.problem) return actual.problem;
  const problems: string[] = [];
  for (const [key, count] of expected.counts) {
    const got = actual.counts.get(key) ?? 0;
    if (got !== count) problems.push(`${key} appears ${got}x, expected ${count}x`);
  }
  for (const key of actual.counts.keys()) {
    if (!expected.counts.has(key)) problems.push(`unexpected ${key}`);
  }
  return problems.length > 0 ? problems.join("; ") : undefined;
}

/** CLDR Vietnamese has the single cardinal (and ordinal) category `other`. */
function checkViPlural(tokens: IcuToken[]): string | undefined {
  for (const token of tokens) {
    if (token.type !== "plural" && token.type !== "select" && token.type !== "selectordinal") {
      continue;
    }
    if (token.type !== "select") {
      const keys = token.cases.map((branch) => branch.key);
      const invalid = keys.filter((key) => key !== "other" && !/^=\d+$/.test(key));
      if (invalid.length > 0) {
        return `{${token.arg}, ${token.type}} uses ${invalid.join(", ")}; Vietnamese allows only "other" and "=N" branches`;
      }
      if (!keys.includes("other")) return `{${token.arg}, ${token.type}} is missing "other"`;
    }
    for (const branch of token.cases) {
      const problem = checkViPlural(branch.tokens);
      if (problem) return problem;
    }
  }
  return undefined;
}

/** Runs every per-message validator on one translation. */
export function validateTranslation(
  msgid: string,
  msgstr: string,
  allowlist: Map<string, string>,
): Finding[] {
  if (msgstr.trim() === "") return [{ rule: "empty", msgid, detail: "msgstr is empty" }];
  const findings: Finding[] = [];
  if (msgstr !== msgstr.normalize("NFC")) {
    findings.push({ rule: "nfc", msgid, detail: "msgstr is not NFC-normalized" });
  }
  try {
    parseIcu(msgstr);
  } catch (error) {
    const detail = error instanceof Error ? error.message.split("\n")[0] : String(error);
    findings.push({ rule: "icu-invalid", msgid, detail });
    return findings;
  }
  const placeholderProblem = comparePlaceholders(msgid, msgstr);
  if (placeholderProblem) {
    findings.push({ rule: "placeholder-parity", msgid, detail: placeholderProblem });
  }
  const tagProblem = compareTags(msgid, msgstr);
  if (tagProblem) findings.push({ rule: "tag-parity", msgid, detail: tagProblem });
  const pluralProblem = checkViPlural(parseIcu(msgstr));
  if (pluralProblem) findings.push({ rule: "vi-plural", msgid, detail: pluralProblem });
  if (msgstr === msgid && !allowlist.has(msgid)) {
    findings.push({
      rule: "english-leftover",
      msgid,
      detail: "msgstr is identical to msgid and not on the identical-ok allowlist",
    });
  }
  return findings;
}

/** Translated and escalated together must cover the shard exactly: no misses, extras, overlap. */
export function validateCoverage(shardMsgids: string[], fragment: Fragment): Finding[] {
  const findings: Finding[] = [];
  const shardSet = new Set(shardMsgids);
  const translated = new Set(Object.keys(fragment.translations));
  const escalated = new Set(fragment.escalated.map((entry) => entry.msgid));
  if (escalated.size !== fragment.escalated.length) {
    findings.push({ rule: "coverage", msgid: "", detail: "duplicate msgids in escalated" });
  }
  for (const entry of fragment.escalated) {
    if (entry.reason.trim() === "") {
      findings.push({ rule: "coverage", msgid: entry.msgid, detail: "escalation without a reason" });
    }
    if (translated.has(entry.msgid)) {
      findings.push({ rule: "coverage", msgid: entry.msgid, detail: "both translated and escalated" });
    }
  }
  for (const msgid of shardMsgids) {
    if (!translated.has(msgid) && !escalated.has(msgid)) {
      findings.push({ rule: "coverage", msgid, detail: "missing from the fragment" });
    }
  }
  for (const msgid of [...translated, ...escalated]) {
    if (!shardSet.has(msgid)) {
      findings.push({ rule: "coverage", msgid, detail: "not in this shard" });
    }
  }
  return findings;
}

/** Validates one shard fragment completely: coverage plus every translation. */
export function validateFragment(
  shardMsgids: string[],
  fragment: Fragment,
  allowlist: Map<string, string>,
): Finding[] {
  const findings = validateCoverage(shardMsgids, fragment);
  const shardSet = new Set(shardMsgids);
  for (const [msgid, msgstr] of Object.entries(fragment.translations)) {
    if (!shardSet.has(msgid)) continue; // already a coverage finding
    findings.push(...validateTranslation(msgid, msgstr, allowlist));
  }
  return findings;
}

// --- File loading ------------------------------------------------------------------------------

/** Parses the allowlist: one `msgid<TAB>reason` per line, `#` comments and blanks skipped. */
export function parseAllowlist(text: string): Map<string, string> {
  const allowlist = new Map<string, string>();
  for (const [index, line] of text.split("\n").entries()) {
    if (line === "" || line.startsWith("#")) continue;
    const tab = line.indexOf("\t");
    if (tab === -1) {
      throw new Error(`Allowlist line ${index + 1} has no tab-separated reason: ${line}`);
    }
    allowlist.set(line.slice(0, tab), line.slice(tab + 1));
  }
  return allowlist;
}

function loadAllowlist(path: string): Map<string, string> {
  return existsSync(path) ? parseAllowlist(readFileSync(path, "utf8")) : new Map();
}

/** Loads and shape-checks a fragment file. Shape errors are findings, not crashes. */
export function loadFragment(path: string): { fragment?: Fragment; findings: Finding[] } {
  const fail = (detail: string): { findings: Finding[] } => ({
    findings: [{ rule: "fragment-shape", msgid: "", detail: `${basename(path)}: ${detail}` }],
  });
  let raw: unknown;
  try {
    raw = JSON.parse(readFileSync(path, "utf8"));
  } catch (error) {
    return fail(error instanceof Error ? error.message : String(error));
  }
  if (typeof raw !== "object" || raw === null) return fail("not a JSON object");
  const { translations, escalated } = raw as Partial<Fragment>;
  if (typeof translations !== "object" || translations === null || Array.isArray(translations)) {
    return fail("`translations` must be an object of msgid -> msgstr");
  }
  for (const [msgid, msgstr] of Object.entries(translations)) {
    if (typeof msgstr !== "string") return fail(`translation for ${JSON.stringify(msgid)} is not a string`);
  }
  if (!Array.isArray(escalated)) return fail("`escalated` must be an array");
  for (const entry of escalated) {
    if (
      typeof entry !== "object" ||
      entry === null ||
      typeof (entry as { msgid?: unknown }).msgid !== "string" ||
      typeof (entry as { reason?: unknown }).reason !== "string"
    ) {
      return fail("every `escalated` entry must be { msgid: string, reason: string }");
    }
  }
  return { fragment: { translations, escalated }, findings: [] };
}

function loadShardMsgids(path: string): string[] {
  const shard = JSON.parse(readFileSync(path, "utf8")) as { messages: { msgid: string }[] };
  return shard.messages.map((message) => message.msgid);
}

// --- CLI ---------------------------------------------------------------------------------------

function printFindings(findings: Finding[]): void {
  for (const finding of findings) {
    const excerpt = finding.msgid.length > 70 ? `${finding.msgid.slice(0, 70)}…` : finding.msgid;
    console.error(`  [${finding.rule}] ${JSON.stringify(excerpt)}: ${finding.detail}`);
  }
}

function checkOne(fragmentPath: string, shardsDir: string, allowlistPath: string): number {
  const shardFile = basename(fragmentPath).replace(/\.vi\.json$/, ".json");
  const shardPath = join(shardsDir, shardFile);
  if (!existsSync(shardPath)) {
    console.error(`No shard file ${shardPath} matches ${fragmentPath}.`);
    return 1;
  }
  const { fragment, findings: shapeFindings } = loadFragment(fragmentPath);
  const findings = fragment
    ? validateFragment(loadShardMsgids(shardPath), fragment, loadAllowlist(allowlistPath))
    : shapeFindings;
  if (findings.length > 0) {
    console.error(`${basename(fragmentPath)}: ${findings.length} finding(s)`);
    printFindings(findings);
    return 1;
  }
  const escalations = fragment ? fragment.escalated.length : 0;
  console.log(
    `${basename(fragmentPath)}: OK -- ${Object.keys(fragment?.translations ?? {}).length} translated, ${escalations} escalated`,
  );
  return 0;
}

function apply(catalogPath: string, shardsDir: string, wavesDir: string, allowlistPath: string): number {
  const allowlist = loadAllowlist(allowlistPath);
  const shardFiles = readdirSync(shardsDir)
    .filter((name) => /^shard-\d+\.json$/.test(name))
    .toSorted();
  if (shardFiles.length === 0) {
    console.error(`No shard files in ${shardsDir}.`);
    return 1;
  }
  const catalog = readPoFile(catalogPath);
  const catalogEntries = new Map(catalog.entries.map((entry) => [entry.msgid, entry]));
  const allFindings: Finding[] = [];
  const merged = new Map<string, string>();
  for (const shardFile of shardFiles) {
    const fragmentPath = join(wavesDir, shardFile.replace(/\.json$/, ".vi.json"));
    if (!existsSync(fragmentPath)) {
      allFindings.push({ rule: "coverage", msgid: "", detail: `missing fragment ${basename(fragmentPath)}` });
      continue;
    }
    const { fragment, findings: shapeFindings } = loadFragment(fragmentPath);
    if (!fragment) {
      allFindings.push(...shapeFindings);
      continue;
    }
    allFindings.push(...validateFragment(loadShardMsgids(join(shardsDir, shardFile)), fragment, allowlist));
    for (const [msgid, msgstr] of Object.entries(fragment.translations)) merged.set(msgid, msgstr);
  }
  // The orchestrator's escalation resolutions: same per-message validators, no shard coverage,
  // and it may not escalate further. Its entries win over shard fragments by applying last.
  const escalationsPath = join(wavesDir, "escalations.vi.json");
  if (existsSync(escalationsPath)) {
    const { fragment, findings: shapeFindings } = loadFragment(escalationsPath);
    if (!fragment) {
      allFindings.push(...shapeFindings);
    } else {
      if (fragment.escalated.length > 0) {
        allFindings.push({
          rule: "fragment-shape",
          msgid: "",
          detail: "escalations.vi.json may not itself escalate",
        });
      }
      for (const [msgid, msgstr] of Object.entries(fragment.translations)) {
        if (!catalogEntries.has(msgid)) {
          allFindings.push({ rule: "coverage", msgid, detail: "escalation resolution for unknown msgid" });
          continue;
        }
        allFindings.push(...validateTranslation(msgid, msgstr, allowlist));
        merged.set(msgid, msgstr);
      }
    }
  }
  if (allFindings.length > 0) {
    console.error(`Refusing to merge: ${allFindings.length} finding(s)`);
    printFindings(allFindings);
    return 1;
  }
  for (const [msgid, msgstr] of merged) {
    const entry = catalogEntries.get(msgid);
    if (!entry) {
      // A shard msgid absent from the catalog means the catalog was re-extracted since sharding.
      console.error(`Shard msgid not in ${catalogPath}: ${JSON.stringify(msgid)} -- re-run shard.ts.`);
      return 1;
    }
    setMsgstr(entry, msgstr);
  }
  writePoFile(catalogPath, catalog);
  const empty = catalog.entries.filter((entry) => entry.msgstr === "");
  console.log(`Merged ${merged.size} translations into ${catalogPath}.`);
  if (empty.length > 0) {
    console.log(`${empty.length} entries remain untranslated (open escalations):`);
    for (const entry of empty) console.log(`  ${JSON.stringify(entry.msgid.slice(0, 70))}`);
  }
  return 0;
}

function audit(catalogPath: string, allowlistPath: string): number {
  const allowlist = loadAllowlist(allowlistPath);
  const catalog: PoCatalog = readPoFile(catalogPath);
  const findings: Finding[] = [];
  for (const entry of catalog.entries) {
    if (entry.msgstr === "") {
      findings.push({ rule: "empty", msgid: entry.msgid, detail: "untranslated" });
    } else if (entry.msgstr === entry.msgid && !allowlist.has(entry.msgid)) {
      findings.push({ rule: "english-leftover", msgid: entry.msgid, detail: "identical to msgid, not allowlisted" });
    }
  }
  if (findings.length > 0) {
    console.error(`${catalogPath}: ${findings.length} finding(s)`);
    printFindings(findings);
    return 1;
  }
  const allowlisted = catalog.entries.filter((entry) => allowlist.has(entry.msgid)).length;
  console.log(
    `${catalogPath}: all ${catalog.entries.length} entries translated (${allowlisted} identical-ok by allowlist).`,
  );
  return 0;
}

function main(): void {
  const { values, positionals } = parseArgs({
    allowPositionals: true,
    options: {
      check: { type: "string" },
      apply: { type: "boolean", default: false },
      audit: { type: "boolean", default: false },
      catalog: { type: "string", default: DEFAULT_CATALOG },
      shards: { type: "string" },
      waves: { type: "string" },
      allowlist: { type: "string", default: DEFAULT_ALLOWLIST },
    },
  });
  const modes = [values.check !== undefined, values.apply, values.audit].filter(Boolean).length;
  if (modes !== 1 || positionals.length > 0) {
    console.error(
      "Usage:\n" +
        "  merge.ts --check <fragment.vi.json> [--shards <dir>] [--allowlist <file>]\n" +
        "  merge.ts --apply --shards <dir> --waves <dir> [--catalog <vi.po>] [--allowlist <file>]\n" +
        "  merge.ts --audit [--catalog <vi.po>] [--allowlist <file>]",
    );
    process.exit(1);
  }
  if (values.check !== undefined) {
    // By convention the plan directory holds `shards/` and `waves/` side by side.
    const shardsDir = values.shards ?? join(dirname(values.check), "..", "shards");
    process.exit(checkOne(values.check, shardsDir, values.allowlist));
  }
  if (values.apply) {
    if (!values.shards || !values.waves) {
      console.error("--apply requires --shards and --waves.");
      process.exit(1);
    }
    process.exit(apply(values.catalog, values.shards, values.waves, values.allowlist));
  }
  process.exit(audit(values.catalog, values.allowlist));
}

if (process.argv[1] === import.meta.filename) main();

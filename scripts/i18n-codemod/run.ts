// The i18n codemod: wraps the user-facing strings `detect.ts` finds, and reports the ones it will
// not touch so a follow-up pass has an exact worklist.
//
//   node scripts/i18n-codemod/run.ts --dry-run --report codemod-residue-report.json
//   node scripts/i18n-codemod/run.ts --apply
//   node scripts/i18n-codemod/run.ts --dry-run --files 'packages/workshop-frontend/src/**/*.tsx'
//
// Nothing is written without `--apply`. `--dry-run` is the default and may be passed explicitly.
import { globSync, readFileSync, writeFileSync } from "node:fs";
import { relative, resolve } from "node:path";
import ts from "typescript6";
import {
  DEFAULT_ALLOWLIST_PATH, REPO_ROOT, SCAN_ROOT, findStringCandidates, isAllowlisted, isExcludedFile,
  listSourceFiles, loadAllowlist, type Allowlist, type DetectOptions, type StringCandidate,
} from "./detect.ts";
import { applyEdits, type PlannedEdit, type ResidueEntry } from "./edits.ts";
import { conflictingBindings, planMacroImports } from "./imports.ts";
import { T_MACRO_SPECIFIER, planAttributeWraps } from "./wrap-attrs.ts";
import { planJsxExpressionWraps } from "./wrap-jsx-expressions.ts";
import { TRANS_MACRO_SPECIFIER, planJsxTextWraps } from "./wrap-jsx-text.ts";

/** What one file's pass produced. */
export interface FileResult {
  /** Repository-relative, forward slashes. */
  file: string;
  /** How many detected strings the file held, allowlist applied. */
  candidates: number;
  edits: PlannedEdit[];
  residue: ResidueEntry[];
  /** The rewritten source, absent when nothing changed. */
  output?: string;
}

/** The report's per-edit record, trimmed to what a reviewer reads. */
interface ReportEdit {
  kind: PlannedEdit["kind"];
  line: number;
  column: number;
  before: string;
  after: string;
}

/** The shape written to `--report`. Consumed by the follow-up pass over the residue. */
interface Report {
  generatedBy: string;
  mode: "dry-run" | "apply";
  scanRoot: string;
  totals: {
    filesScanned: number;
    filesChanged: number;
    candidates: number;
    edits: number;
    editsByKind: Record<string, number>;
    residue: number;
    residueByReason: Record<string, number>;
    residueFiles: number;
    residueGroups: number;
  };
  files: {
    file: string;
    edits: ReportEdit[];
    residue: ResidueEntry[];
  }[];
}

const REPORT_SOURCE = "scripts/i18n-codemod/run.ts";

/** Options for {@link transformFile}. */
export interface TransformOptions {
  allowlist?: Allowlist;
  /**
   * What the detector looks for. Passed straight through, so the codemod and the guard can never
   * disagree about what counts as a user-facing string.
   */
  detect?: DetectOptions;
}

function parse(file: string, source: string): ts.SourceFile {
  return ts.createSourceFile(file, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
}

function syntaxErrorsIn(file: string, source: string): string[] {
  const { diagnostics } = ts.transpileModule(source, {
    fileName: file,
    reportDiagnostics: true,
    compilerOptions: { jsx: ts.JsxEmit.Preserve, target: ts.ScriptTarget.Latest },
  });
  return (diagnostics ?? []).map((diagnostic) =>
    ts.flattenDiagnosticMessageText(diagnostic.messageText, " "));
}

function refuse(
  candidates: readonly StringCandidate[], reason: ResidueEntry["reason"],
): ResidueEntry[] {
  return candidates.map((candidate): ResidueEntry => ({
    kind: candidate.kind,
    attribute: candidate.attribute,
    parentTag: candidate.parentTag,
    text: candidate.text,
    line: candidate.line,
    column: candidate.column,
    reason,
  }));
}

/**
 * Plans and, if anything is planned, produces the rewritten source for one file.
 *
 * The rewrite is verified two ways. Its output must re-parse -- a file that does not is discarded
 * whole and every candidate in it becomes residue, because a partially-applied codemod is worse
 * than one that declined. And the residue is then re-derived from that output, because every wrap
 * and every added import moves the text below it: a residue entry has to point at the line a reader
 * will actually open, not the line the string used to be on.
 */
export function transformFile(
  file: string, source: string, options: TransformOptions = {},
): FileResult {
  // Every transform here rewrites JSX, and `detect.ts` reports a different kind of candidate for a
  // `.ts` file, which nothing below would plan an edit for. Refusing is clearer than silently
  // counting candidates the codemod cannot act on.
  if (!file.endsWith(".tsx")) {
    throw new Error(`The codemod transforms .tsx files only, not ${file}.`);
  }

  const planned = planFile(file, source, options);
  if (planned.output === undefined) return planned;

  // Idempotency makes this exact rather than approximate: the second pass finds the same strings,
  // refuses them for the same reasons, and plans nothing. If it ever planned something, the
  // codemod would not be idempotent and the report could not be trusted either.
  const reprojected = planFile(file, planned.output, options);
  if (reprojected.edits.length > 0 || reprojected.residue.length !== planned.residue.length) {
    throw new Error(
      `Rewriting ${file} was not idempotent: a second pass planned ${reprojected.edits.length} ` +
      `edit(s) and found ${reprojected.residue.length} residue entries, not ${planned.residue.length}.`,
    );
  }

  return { ...planned, residue: reprojected.residue };
}

function planFile(file: string, source: string, options: TransformOptions): FileResult {
  const sourceFile = parse(file, source);
  const candidates = findStringCandidates(file, source, options.detect)
    .filter((candidate) => !options.allowlist || !isAllowlisted(options.allowlist, candidate));

  const count = candidates.length;
  // Per macro, because a file that already uses `t` as a local can still take `<Trans>`.
  const conflicts = new Set(conflictingBindings(sourceFile));
  const tConflict = conflicts.has(T_MACRO_SPECIFIER);
  const text = conflicts.has(TRANS_MACRO_SPECIFIER)
    ? { edits: [], residue: refuse(byKind(candidates, "jsx-text"), "identifier-conflict") }
    : planJsxTextWraps(source, sourceFile, candidates);
  const attributes = tConflict
    ? { edits: [], residue: refuse(byKind(candidates, "jsx-attribute"), "identifier-conflict") }
    : planAttributeWraps(source, sourceFile, candidates);
  const expressions = tConflict
    ? { edits: [], residue: refuse(byKind(candidates, "jsx-expression"), "identifier-conflict") }
    : planJsxExpressionWraps(source, sourceFile, candidates);

  const imports = planMacroImports(source, sourceFile, {
    trans: text.edits.length > 0,
    t: attributes.edits.length > 0 || expressions.edits.length > 0,
  });
  const edits = [...text.edits, ...attributes.edits, ...expressions.edits, ...imports.edits];
  const residue = [...text.residue, ...attributes.residue, ...expressions.residue];

  if (edits.length === 0) return { file, candidates: count, edits, residue };

  const output = applyEdits(source, edits);
  const errors = syntaxErrorsIn(file, output);
  if (errors.length > 0) {
    return {
      file, candidates: count, edits: [],
      residue: refuse(candidates, "output-parse-failed"),
    };
  }

  return { file, candidates: count, edits, residue, output };
}

function byKind(
  candidates: readonly StringCandidate[], kind: StringCandidate["kind"],
): StringCandidate[] {
  return candidates.filter((candidate) => candidate.kind === kind);
}

/** The command line, after parsing. */
interface Options {
  apply: boolean;
  files?: string;
  directory: string;
  reportPath?: string;
  allowlistPath: string;
}

function parseArguments(argv: string[]): Options {
  const options: Options = {
    apply: false,
    directory: SCAN_ROOT,
    allowlistPath: DEFAULT_ALLOWLIST_PATH,
  };

  for (let index = 0; index < argv.length; index++) {
    const argument = argv[index];
    if (argument === "--dry-run") continue;
    if (argument === "--apply") {
      options.apply = true;
    } else if (argument === "--files" || argument === "--dir" || argument === "--report" ||
      argument === "--allowlist") {
      const value = argv[++index];
      if (!value) throw new Error(`Missing value for ${argument}.`);
      if (argument === "--files") options.files = value;
      else if (argument === "--dir") options.directory = resolve(REPO_ROOT, value);
      else if (argument === "--report") options.reportPath = resolve(REPO_ROOT, value);
      else options.allowlistPath = resolve(REPO_ROOT, value);
    } else {
      throw new Error(`Unknown argument: ${argument}`);
    }
  }

  return options;
}

function targetFiles(options: Options): string[] {
  // `.tsx` only: every transform here is a JSX rewrite. The guard reports on `.ts` files too, but
  // wrapping those is hand work this codemod does not know how to do.
  if (!options.files) return listSourceFiles(options.directory, [".tsx"]);
  return globSync(options.files, { cwd: REPO_ROOT })
    .map((match) => resolve(REPO_ROOT, match))
    .filter((path) => path.endsWith(".tsx") && !isExcludedFile(path))
    .toSorted((a, b) => (a < b ? -1 : a > b ? 1 : 0));
}

function countBy<T extends string>(values: readonly T[]): Record<string, number> {
  const counts: Record<string, number> = {};
  for (const value of values.toSorted()) counts[value] = (counts[value] ?? 0) + 1;
  return counts;
}

function buildReport(results: readonly FileResult[], options: Options): Report {
  const edits = results.flatMap((result) => result.edits);
  const residue = results.flatMap((result) => result.residue);
  const groups = new Set(residue.map((entry) => entry.group).filter((group) => group !== undefined));

  return {
    generatedBy: REPORT_SOURCE,
    mode: options.apply ? "apply" : "dry-run",
    scanRoot: relative(REPO_ROOT, options.directory).replaceAll("\\", "/"),
    totals: {
      filesScanned: results.length,
      filesChanged: results.filter((result) => result.output !== undefined).length,
      candidates: results.reduce((total, result) => total + result.candidates, 0),
      edits: edits.length,
      editsByKind: countBy(edits.map((edit) => edit.kind)),
      residue: residue.length,
      residueByReason: countBy(residue.map((entry) => entry.reason)),
      residueFiles: results.filter((result) => result.residue.length > 0).length,
      residueGroups: groups.size,
    },
    files: results
      .filter((result) => result.edits.length > 0 || result.residue.length > 0)
      .map((result) => ({
        file: result.file,
        edits: result.edits.map((edit) => ({
          kind: edit.kind,
          line: edit.line,
          column: edit.column,
          before: edit.before,
          after: edit.replacement,
        })),
        residue: result.residue,
      })),
  };
}

function main(): void {
  let options: Options;
  try {
    options = parseArguments(process.argv.slice(2));
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 2;
    return;
  }

  const allowlist = loadAllowlist(options.allowlistPath);
  const results: FileResult[] = [];

  for (const path of targetFiles(options)) {
    const file = relative(REPO_ROOT, path).replaceAll("\\", "/");
    let source: string;
    try {
      source = readFileSync(path, "utf8");
    } catch (error) {
      console.error(`Could not read ${file}:`, error);
      process.exitCode = 2;
      return;
    }

    let result: FileResult;
    try {
      result = transformFile(file, source, { allowlist });
    } catch (error) {
      console.error(`Could not transform ${file}:`, error);
      process.exitCode = 2;
      return;
    }
    results.push(result);

    if (options.apply && result.output !== undefined) {
      try {
        writeFileSync(path, result.output);
      } catch (error) {
        console.error(`Could not write ${file}:`, error);
        process.exitCode = 2;
        return;
      }
    }
  }

  const report = buildReport(results, options);
  if (options.reportPath) {
    try {
      writeFileSync(options.reportPath, `${JSON.stringify(report, undefined, 2)}\n`);
    } catch (error) {
      console.error(`Could not write the report:`, error);
      process.exitCode = 2;
      return;
    }
  }

  const { totals } = report;
  console.log(
    `${options.apply ? "Applied" : "Planned"} ${totals.edits} edit(s) in ` +
    `${totals.filesChanged} of ${totals.filesScanned} file(s), from ${totals.candidates} ` +
    `candidate(s).`,
  );
  console.log(`  by kind:   ${JSON.stringify(totals.editsByKind)}`);
  console.log(
    `  residue:   ${totals.residue} in ${totals.residueFiles} file(s), ` +
    `${totals.residueGroups} group(s)`,
  );
  console.log(`  by reason: ${JSON.stringify(totals.residueByReason)}`);
  if (options.reportPath) {
    console.log(`  report:    ${relative(REPO_ROOT, options.reportPath)}`);
  }
}

if (process.argv[1] && import.meta.filename === resolve(process.argv[1])) main();

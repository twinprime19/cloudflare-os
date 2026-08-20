// Reports user-facing strings in the frontend that no i18n macro wraps yet. A thin front end for
// `scripts/i18n-codemod/detect.ts`: every rule about what counts lives there, shared with the
// codemod, so this file only chooses what to scan and how to print it.
//
// Exit code 1 when anything is reported, so it can gate a lint run once the extraction is done.
//
//   node scripts/check-unlocalized-strings.ts [--dir <path>] [--allowlist <path>] [--json]
//                                             [--helper-literals]
//
// `--dir` is resolved against the repository root and defaults to the frontend's `src`.
// JSX text, whitelisted attributes, JSX expression containers and `.ts` string literals are all
// reported by default, because the codemod transforms all four. `--helper-literals` additionally
// reports literals a `.tsx` file builds outside JSX; see `DetectOptions.includeHelperLiterals`.
import { relative, resolve } from "node:path";
import {
  DEFAULT_ALLOWLIST_PATH, REPO_ROOT, SCAN_ROOT, loadAllowlist, scanDirectory,
  type StringCandidate,
} from "./i18n-codemod/detect.ts";

/** The command line, after parsing. */
interface Options {
  directory: string;
  allowlistPath: string;
  json: boolean;
  includeHelperLiterals: boolean;
}

function parseArguments(argv: string[]): Options {
  const options: Options = {
    directory: SCAN_ROOT,
    allowlistPath: DEFAULT_ALLOWLIST_PATH,
    json: false,
    includeHelperLiterals: false,
  };

  for (let index = 0; index < argv.length; index++) {
    const argument = argv[index];
    if (argument === "--json") {
      options.json = true;
    } else if (argument === "--helper-literals") {
      options.includeHelperLiterals = true;
    } else if (argument === "--dir" || argument === "--allowlist") {
      const value = argv[++index];
      if (!value) throw new Error(`Missing value for ${argument}.`);
      if (argument === "--dir") options.directory = resolve(REPO_ROOT, value);
      else options.allowlistPath = resolve(REPO_ROOT, value);
    } else {
      throw new Error(`Unknown argument: ${argument}`);
    }
  }

  return options;
}

function describe(candidate: StringCandidate): string {
  const where = `${candidate.file}:${candidate.line}:${candidate.column}`;
  const what = candidate.kind === "jsx-attribute" ? `${candidate.attribute}=` : "";
  return `${where}  ${what}${JSON.stringify(candidate.text)}`;
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

  let candidates: StringCandidate[];
  try {
    candidates = scanDirectory(options.directory, {
      allowlist: loadAllowlist(options.allowlistPath),
      includeHelperLiterals: options.includeHelperLiterals,
    });
  } catch (error) {
    console.error(`Could not scan ${relative(REPO_ROOT, options.directory)}:`, error);
    process.exitCode = 2;
    return;
  }

  if (options.json) {
    console.log(JSON.stringify(candidates, undefined, 2));
  } else {
    for (const candidate of candidates) console.log(describe(candidate));
  }

  if (candidates.length === 0) {
    if (!options.json) console.log("No unlocalized user-facing strings found.");
    return;
  }

  if (!options.json) {
    const count = (kind: StringCandidate["kind"]): number =>
      candidates.filter((candidate) => candidate.kind === kind).length;
    console.error(
      `\n${candidates.length} unlocalized string(s): ` +
      `${count("jsx-text")} in JSX text, ${count("jsx-attribute")} in attributes, ` +
      `${count("string-literal")} in .ts string literals, ` +
      `${count("jsx-expression")} in JSX expressions.\n` +
      "Wrap them with `<Trans>` or `` t`...` ``, or record an intentional exception in " +
      `${relative(REPO_ROOT, options.allowlistPath)}.`,
    );
  }
  process.exitCode = 1;
}

main();

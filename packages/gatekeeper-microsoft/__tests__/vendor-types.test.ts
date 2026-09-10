// The vendor serves every resource type's `.d.ts` as one string, joined with a newline
// (`GatekeeperVendor.getTypeScriptTypes`). Concatenating two independently written declaration
// files is only safe if they were written not to collide, and nothing in the type system enforces
// that: `tsc` checks each file on its own, and the joined text is only ever assembled at runtime,
// where a duplicate identifier would surface as a parse failure in the Workshop's type database
// rather than at build time. This parses what the vendor actually serves.
//
// The files are read from disk rather than imported: under vitest a `.txt` import resolves to a
// module reference, not the file's contents. `types.txt` and `teams-types.txt`, which the vendor
// imports, are symlinks to the `.d.ts` files read here, so this is the same text — but the join is
// reproduced here rather than called, so it has to be kept in step with the vendor's.

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
// The JS-based compiler the repo's build scripts use. The workspace `typescript` is 7.x (tsgo),
// which ships version metadata only and no parser API.
import ts from "typescript6";

const SOURCES = ["types.d.ts", "teams-types.d.ts"];

function read(name: string): string {
  return readFileSync(fileURLToPath(new URL(`../src/${name}`, import.meta.url)), "utf8");
}

/** The same sources, in the same order, joined the same way `GatekeeperVendor` joins them. */
const joined = SOURCES.map(read).join("\n");

function parse(text: string, fileName: string): ts.SourceFile {
  return ts.createSourceFile(fileName, text, ts.ScriptTarget.ES2022, false, ts.ScriptKind.TS);
}

/** Every name the file introduces at the top level, including re-exported ones. */
function declaredNames(source: ts.SourceFile): string[] {
  const names: string[] = [];
  for (const statement of source.statements) {
    if ((ts.isTypeAliasDeclaration(statement) || ts.isInterfaceDeclaration(statement)
        || ts.isClassDeclaration(statement) || ts.isFunctionDeclaration(statement))
        && statement.name) {
      names.push(statement.name.text);
    } else if (ts.isVariableStatement(statement)) {
      for (const declaration of statement.declarationList.declarations) {
        if (ts.isIdentifier(declaration.name)) names.push(declaration.name.text);
      }
    } else if (ts.isImportDeclaration(statement)) {
      const clause = statement.importClause;
      if (clause?.name) names.push(clause.name.text);
      if (clause?.namedBindings && ts.isNamedImports(clause.namedBindings)) {
        for (const element of clause.namedBindings.elements) names.push(element.name.text);
      }
    } else if (ts.isExportDeclaration(statement) && statement.exportClause
        && ts.isNamedExports(statement.exportClause) && !statement.moduleSpecifier) {
      // `export type { Cursor }` re-exports a name already bound in this file, so it introduces
      // nothing new; a `from` clause would.
      continue;
    }
  }
  return names;
}

describe("vendor TypeScript types", () => {
  it("parses as one module with no syntax errors", () => {
    const source = parse(joined, "microsoft-types.d.ts");
    const diagnostics = (source as ts.SourceFile & { parseDiagnostics?: ts.Diagnostic[] })
        .parseDiagnostics ?? [];

    expect(diagnostics.map(
        diagnostic => ts.flattenDiagnosticMessageText(diagnostic.messageText, " "))).toEqual([]);
  });

  it("declares every top-level name exactly once", () => {
    // The collision this guards against is real and specific: the mail types import and re-export
    // `Cursor`, so the Teams types reach it through an inline import instead. A second import would
    // leave the joined text with a duplicate identifier.
    const names = declaredNames(parse(joined, "microsoft-types.d.ts"));
    const duplicates = names.filter((name, index) => names.indexOf(name) !== index);

    expect(duplicates).toEqual([]);
  });

  it("still serves each resource's own types, which the per-resource gatekeepers return", () => {
    // Each file has to stand on its own too: a `Gatekeeper.getTypeScriptTypes()` serves one of
    // them alone, so neither may depend on a name the other happens to bring.
    for (const name of SOURCES) {
      const source = parse(read(name), name);
      const diagnostics = (source as ts.SourceFile & { parseDiagnostics?: ts.Diagnostic[] })
          .parseDiagnostics ?? [];

      expect(diagnostics).toEqual([]);
    }
    expect(read("teams-types.d.ts")).toContain("export interface TeamsSession");
    expect(read("types.d.ts")).toContain("export interface OutlookMailSession");
  });

  it("names the types the resource descriptions point at", () => {
    // `ResourceDescription.tsType` must be an export of the served types.
    expect(joined).toContain("export interface OutlookMailSession");
    expect(joined).toContain("export interface TeamsSession");
  });
});

// Adds the Lingui macro imports a wrapped file needs, in the shape the hand-wrapped reference file
// (`packages/workshop-frontend/src/components/AutoApproveConfirmDialog.tsx`) uses: named imports
// from `@lingui/core/macro` and `@lingui/react/macro`, placed after the package imports and before
// the relative ones, in that order.
//
// Idempotent by construction. An import that is already there produces no edit, which is what makes
// re-running the codemod a no-op.
import ts from "typescript6";
import { emptyPlan, type Plan, type PlannedEdit } from "./edits.ts";
import { T_MACRO_MODULE, T_MACRO_SPECIFIER } from "./wrap-attrs.ts";
import { TRANS_MACRO_MODULE, TRANS_MACRO_SPECIFIER } from "./wrap-jsx-text.ts";

/** Which macros the rewritten file will use. */
export interface MacroImportNeeds {
  /** `<Trans>` was inserted somewhere. */
  trans: boolean;
  /** `` t`...` `` was inserted somewhere. */
  t: boolean;
}

/** How the file writes its own imports, so an added one does not look bolted on. */
export interface ImportStyle {
  quote: string;
  semicolon: boolean;
}

// Emitted in the order the reference file uses: the core macro, then the React one.
const MACROS: readonly { specifier: string; module: string; need: keyof MacroImportNeeds }[] = [
  { specifier: T_MACRO_SPECIFIER, module: T_MACRO_MODULE, need: "t" },
  { specifier: TRANS_MACRO_SPECIFIER, module: TRANS_MACRO_MODULE, need: "trans" },
];

function importDeclarations(sourceFile: ts.SourceFile): ts.ImportDeclaration[] {
  return sourceFile.statements.filter((statement) => ts.isImportDeclaration(statement));
}

function moduleNameOf(declaration: ts.ImportDeclaration): string | undefined {
  const specifier = declaration.moduleSpecifier;
  return ts.isStringLiteral(specifier) ? specifier.text : undefined;
}

function namedImportsOf(declaration: ts.ImportDeclaration): ts.NamedImports | undefined {
  const bindings = declaration.importClause?.namedBindings;
  return bindings && ts.isNamedImports(bindings) ? bindings : undefined;
}

// Anything that introduces a name into a scope in the file. A declaration whose name is a
// destructuring pattern is reached through its binding elements, so the pattern itself is not here.
function declaredNameOf(node: ts.Node): string | undefined {
  if (
    ts.isVariableDeclaration(node) || ts.isParameter(node) || ts.isBindingElement(node) ||
    ts.isFunctionDeclaration(node) || ts.isClassDeclaration(node) ||
    ts.isImportSpecifier(node) || ts.isNamespaceImport(node) || ts.isImportClause(node)
  ) {
    const name: ts.Node | undefined = node.name;
    return name && ts.isIdentifier(name) ? name.text : undefined;
  }
  return undefined;
}

function isMacroImportOf(node: ts.Node, module: string): boolean {
  if (!ts.isImportSpecifier(node)) return false;
  const declaration = node.parent.parent.parent;
  return ts.isImportDeclaration(declaration) && moduleNameOf(declaration) === module;
}

/**
 * The macro names this file already uses for something else.
 *
 * A macro is resolved by its binding, so any other declaration of `t` or `Trans` -- another import,
 * a variable, a callback parameter -- either collides with the import the codemod would add or
 * silently shadows it at a call site. `t` in particular is a common name for a timeout handle or a
 * `map` parameter. Neither case is worth guessing at, so the affected macro is not used in that
 * file. The two are reported separately because a file that shadows `t` can still take `<Trans>`.
 */
export function conflictingBindings(sourceFile: ts.SourceFile): string[] {
  const conflicts = new Set<string>();

  const visit = (node: ts.Node): void => {
    const name = declaredNameOf(node);
    const macro = name === undefined
      ? undefined
      : MACROS.find((entry) => entry.specifier === name);
    // An existing import of the macro itself is the idempotent case, not a conflict.
    if (macro && !isMacroImportOf(node, macro.module)) conflicts.add(macro.specifier);
    ts.forEachChild(node, visit);
  };
  visit(sourceFile);

  return [...conflicts].toSorted();
}

/** The quote character and semicolon habit of the file's last import, or the frontend default. */
export function detectImportStyle(source: string, sourceFile: ts.SourceFile): ImportStyle {
  const declarations = importDeclarations(sourceFile);
  const last = declarations.at(-1);
  if (!last) return { quote: "'", semicolon: false };
  return {
    quote: source[last.moduleSpecifier.getStart(sourceFile)] ?? "'",
    semicolon: source[last.end - 1] === ";",
  };
}

function declarationText(specifier: string, module: string, style: ImportStyle): string {
  return `import { ${specifier} } from ${style.quote}${module}${style.quote}` +
    (style.semicolon ? ";" : "");
}

/**
 * Where a new import declaration goes: after the last package import, so it lands above the
 * relative ones. Falls back to the last import, then to the top of the file.
 */
function insertionOffset(sourceFile: ts.SourceFile): number {
  const declarations = importDeclarations(sourceFile);
  const packageImports = declarations.filter((declaration) => {
    const module = moduleNameOf(declaration);
    return module !== undefined && !module.startsWith(".") && !module.startsWith("/");
  });
  const anchor = packageImports.at(-1) ?? declarations.at(-1);
  if (anchor) return anchor.end;
  // No imports at all: above the first statement but below any comment that introduces the file.
  return sourceFile.statements[0]?.getStart(sourceFile) ?? 0;
}

/**
 * Plans the macro imports for a file that is about to be wrapped.
 *
 * Returns no edits when the imports are already present, which is the whole of the codemod's
 * idempotency on the import side. When a needed name is bound to something else the file is
 * reported through {@link conflictingBindings} rather than edited, so callers must check that
 * first.
 */
export function planMacroImports(
  source: string,
  sourceFile: ts.SourceFile,
  needs: MacroImportNeeds,
): Plan {
  const plan = emptyPlan();
  const style = detectImportStyle(source, sourceFile);
  const declarations = importDeclarations(sourceFile);
  const fresh: string[] = [];

  for (const macro of MACROS) {
    if (!needs[macro.need]) continue;

    const existing = declarations.find((declaration) =>
      moduleNameOf(declaration) === macro.module && declaration.importClause?.isTypeOnly !== true);

    if (!existing) {
      fresh.push(declarationText(macro.specifier, macro.module, style));
      continue;
    }

    const named = namedImportsOf(existing);
    if (named?.elements.some((element) => element.name.text === macro.specifier)) continue;

    const last = named?.elements.at(-1);
    if (last) {
      plan.edits.push(insertion(source, sourceFile, last.end, `, ${macro.specifier}`));
    } else {
      // A side-effect or default-only import of the same module: leave it and add our own.
      fresh.push(declarationText(macro.specifier, macro.module, style));
    }
  }

  if (fresh.length > 0) {
    const offset = insertionOffset(sourceFile);
    const text = declarations.length > 0
      ? `\n${fresh.join("\n")}`
      : `${fresh.join("\n")}\n\n`;
    plan.edits.push(insertion(source, sourceFile, offset, text));
  }

  return plan;
}

function insertion(
  source: string, sourceFile: ts.SourceFile, offset: number, text: string,
): PlannedEdit {
  const { line, character } = sourceFile.getLineAndCharacterOfPosition(offset);
  return {
    kind: "import",
    start: offset,
    end: offset,
    replacement: text,
    line: line + 1,
    column: character + 1,
    before: "",
  };
}

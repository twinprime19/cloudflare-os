// Retypes a whitelisted attribute's string value as a `` t`...` `` macro call.
//
// `t` from `@lingui/core/macro` compiles to `i18n._(...)` against the global instance, so this
// needs no hook and works in any component, class or function. The trade is that the value is read
// at render time from whatever locale is active -- correct for a client SPA, and this app is one.
import ts from "typescript6";
import type { StringCandidate } from "./detect.ts";
import { emptyPlan, type Plan, type ResidueEntry } from "./edits.ts";

/** The macro import a wrap needs. */
export const T_MACRO_SPECIFIER = "t";

/** Where {@link T_MACRO_SPECIFIER} comes from. */
export const T_MACRO_MODULE = "@lingui/core/macro";

// A template literal reads syntax a JSX string attribute does not, and a JSX string attribute
// decodes HTML entities a template literal would keep verbatim. Either way the message would stop
// being the string the source says it is, so the value is left alone.
const UNSAFE_IN_TEMPLATE = /[\n\r\\`{}&]/;

function attributeIndex(source: ts.SourceFile): Map<number, ts.JsxAttribute> {
  const index = new Map<number, ts.JsxAttribute>();
  const visit = (node: ts.Node): void => {
    if (ts.isJsxAttribute(node)) {
      const literal = literalOf(node);
      if (literal) index.set(literal.getStart(source), node);
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
  return index;
}

function literalOf(node: ts.JsxAttribute): ts.StringLiteral | undefined {
  const initializer = node.initializer;
  if (!initializer) return undefined;
  if (ts.isStringLiteral(initializer)) return initializer;
  if (
    ts.isJsxExpression(initializer) && initializer.expression &&
    ts.isStringLiteral(initializer.expression)
  ) {
    return initializer.expression;
  }
  return undefined;
}

function residueOf(candidate: StringCandidate): ResidueEntry {
  return {
    kind: candidate.kind,
    attribute: candidate.attribute,
    text: candidate.text,
    line: candidate.line,
    column: candidate.column,
    reason: "attribute-unsafe-value",
  };
}

/**
 * Plans a `` t`...` `` rewrite for every attribute candidate whose value survives the move into a
 * template literal.
 *
 * `source` must be the text `candidates` were detected in, and `sourceFile` its parse. Candidates
 * that are not `jsx-attribute` are ignored.
 */
export function planAttributeWraps(
  source: string,
  sourceFile: ts.SourceFile,
  candidates: readonly StringCandidate[],
): Plan {
  const plan = emptyPlan();
  const index = attributeIndex(sourceFile);

  for (const candidate of candidates) {
    if (candidate.kind !== "jsx-attribute") continue;

    const attribute = index.get(candidate.start);
    if (!attribute) {
      throw new Error(`No JSX attribute at offset ${candidate.start} in ${candidate.file}.`);
    }

    const value = source.slice(candidate.textStart, candidate.textEnd);
    if (UNSAFE_IN_TEMPLATE.test(value)) {
      plan.residue.push(residueOf(candidate));
      continue;
    }

    // The span is the string literal with its quotes either way. A braced value already sits in
    // the expression container the macro call needs; a bare one has to grow a pair of braces.
    const braced = attribute.initializer !== undefined &&
      ts.isJsxExpression(attribute.initializer);
    const call = `${T_MACRO_SPECIFIER}\`${value}\``;

    plan.edits.push({
      kind: "jsx-attribute",
      start: candidate.start,
      end: candidate.end,
      replacement: braced ? call : `{${call}}`,
      line: candidate.line,
      column: candidate.column,
      before: source.slice(candidate.start, candidate.end),
    });
  }

  return plan;
}

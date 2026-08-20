// Wraps a JSX text node in `<Trans>`, but only when the whole message is in that one node.
//
// The wrap replaces exactly the span `detect.ts` reported, so the children of `<Trans>` are the
// source bytes unchanged -- indentation and line breaks included, because JSX renders them and
// Lingui's macro collapses them the same way React does. Anything the transform cannot prove is a
// self-contained message is reported instead of guessed at.
import ts from "typescript6";
import type { StringCandidate } from "./detect.ts";
import { emptyPlan, type Plan, type PlannedEdit, type ResidueEntry, type ResidueReason } from "./edits.ts";

/** The macro import a wrap needs. */
export const TRANS_MACRO_SPECIFIER = "Trans";

/** Where {@link TRANS_MACRO_SPECIFIER} comes from. */
export const TRANS_MACRO_MODULE = "@lingui/react/macro";

// A `<Trans>` wrap turns its children into one message. `{` and `}` in JSX text are legal but read
// as ICU syntax to the extractor, and `>` closes nothing but is ambiguous enough to leave alone.
const UNSUPPORTED_TEXT = /[<>{}]/;

type JsxParent = ts.JsxElement | ts.JsxFragment;

function isBlankText(node: ts.Node): boolean {
  return ts.isJsxText(node) && node.text.trim() === "";
}

function contentChildren(parent: JsxParent): ts.JsxChild[] {
  return parent.children.filter((child) => !isBlankText(child));
}

/**
 * Why this text node cannot be wrapped on its own, or `undefined` if it can.
 *
 * The rule is that the text must be the whole message. Sibling expressions and sibling elements
 * with children both mean it is only part of one, and the pieces have to move into a single
 * `<Trans>` together -- a rewrite that needs to understand the sentence, not just its offsets.
 */
function refuseReason(node: ts.JsxText, parent: JsxParent): ResidueReason | undefined {
  if (UNSUPPORTED_TEXT.test(node.text)) return "unsupported-text-content";

  const siblings = contentChildren(parent).filter((child) => child !== node);
  if (siblings.length === 0) return undefined;

  if (siblings.some((child) => ts.isJsxExpression(child) && child.expression !== undefined)) {
    return "interpolated";
  }
  // Two runs of text with something between them: whatever that something is, it splits a sentence.
  if (siblings.some((child) => ts.isJsxText(child))) return "split-by-markup";
  // Every other sibling is self-closing, and this is the only run of text, so nothing else in the
  // element renders words: an icon beside a label. The label is the whole message.
  if (siblings.every((child) => ts.isJsxSelfClosingElement(child))) return undefined;
  // An element with children of its own does render text, and that text is part of this sentence.
  return "split-by-markup";
}

function jsxTextIndex(source: ts.SourceFile): Map<number, ts.JsxText> {
  const index = new Map<number, ts.JsxText>();
  const visit = (node: ts.Node): void => {
    if (ts.isJsxText(node)) index.set(node.pos, node);
    ts.forEachChild(node, visit);
  };
  visit(source);
  return index;
}

function parentOf(node: ts.JsxText): JsxParent | undefined {
  const parent = node.parent;
  if (!parent) return undefined;
  return ts.isJsxElement(parent) || ts.isJsxFragment(parent) ? parent : undefined;
}

function residueOf(
  candidate: StringCandidate, reason: ResidueReason, parent?: JsxParent,
): ResidueEntry {
  return {
    kind: candidate.kind,
    parentTag: candidate.parentTag,
    text: candidate.text,
    line: candidate.line,
    column: candidate.column,
    reason,
    group: parent ? `${candidate.file}:${parent.pos}` : undefined,
  };
}

/**
 * Plans a `<Trans>` wrap for every JSX text candidate that is a message on its own.
 *
 * `source` must be the text `candidates` were detected in, and `sourceFile` its parse, so that the
 * offsets line up. Candidates that are not `jsx-text` are ignored.
 */
export function planJsxTextWraps(
  source: string,
  sourceFile: ts.SourceFile,
  candidates: readonly StringCandidate[],
): Plan {
  const plan = emptyPlan();
  const index = jsxTextIndex(sourceFile);

  for (const candidate of candidates) {
    if (candidate.kind !== "jsx-text") continue;

    const node = index.get(candidate.start);
    const parent = node && parentOf(node);
    if (!node || !parent) {
      // A candidate with no matching node means the caller mismatched source and parse.
      throw new Error(`No JSX text node at offset ${candidate.start} in ${candidate.file}.`);
    }

    const reason = refuseReason(node, parent);
    if (reason) {
      plan.residue.push(residueOf(candidate, reason, parent));
      continue;
    }

    plan.edits.push(wrapEdit(source, candidate));
  }

  return plan;
}

function wrapEdit(source: string, candidate: StringCandidate): PlannedEdit {
  const before = source.slice(candidate.textStart, candidate.textEnd);
  return {
    kind: "jsx-text",
    start: candidate.textStart,
    end: candidate.textEnd,
    // `before` verbatim: the bytes between the tags are what the user sees.
    replacement: `<Trans>${before}</Trans>`,
    line: candidate.line,
    column: candidate.column,
    before,
  };
}

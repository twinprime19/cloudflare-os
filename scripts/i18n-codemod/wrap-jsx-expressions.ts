// Retypes a string literal that sits inside a JSX expression container as a `` t`...` `` macro call:
// `{busy ? 'Saving…' : 'Save'}` becomes ``{busy ? t`Saving…` : t`Save`}``.
//
// The literal keeps its place in the expression, so only the literal's own span is replaced and
// everything around it -- the condition, the operators, the formatting -- is untouched. That is the
// same splice-a-span discipline the other two planners use, for the same reason: the guarantee is
// that nothing but the wrap changes.
//
// The hard part is not the rewrite, it is knowing when the literal is a whole message. A branch of a
// ternary usually is; a piece of a concatenation, a `${...}` inside a bigger template, or a branch
// whose sentence continues in the sibling JSX text is not, and neither is a ternary that is really a
// plural. Those are refused and reported. Over-wrapping is worse than residue here, because a
// fragment in the catalog is invisible to every automated check and only a translator ever sees it.
import ts from "typescript6";
import { hasLetters, type StringCandidate } from "./detect.ts";
import { emptyPlan, type Plan, type PlannedEdit, type ResidueEntry, type ResidueReason } from "./edits.ts";
// The same macro the attribute planner uses, imported rather than restated so there is one name for
// it and one module it comes from.
import { T_MACRO_SPECIFIER } from "./wrap-attrs.ts";

// Characters that would not survive the move into a template literal. Tested against the *cooked*
// value, which is what the message will say: `'It\'s'` cooks to `It's` and is perfectly safe, while
// `'a\nb'` cooks to a real newline and `'${'` would open a substitution. A backtick or a backslash
// would need escaping, and escaping is exactly the kind of quiet rewriting this codemod avoids.
// Braces are refused because the extractor reads them as ICU syntax.
//
// This is deliberately not the attribute planner's predicate. A JSX attribute string decodes HTML
// entities, so `&` matters there; a JavaScript string literal does not, so it does not matter here.
const UNSAFE_IN_TEMPLATE = /[\n\r\\`{}]/;

type StringNode = ts.StringLiteral | ts.NoSubstitutionTemplateLiteral | ts.TemplateExpression;

function isStringNode(node: ts.Node): node is StringNode {
  return ts.isStringLiteral(node) || ts.isNoSubstitutionTemplateLiteral(node) ||
    ts.isTemplateExpression(node);
}

function unwrap(node: ts.Expression): ts.Expression {
  let current = node;
  while (ts.isParenthesizedExpression(current) || ts.isNonNullExpression(current)) {
    current = current.expression;
  }
  return current;
}

/**
 * Whether an expression can only ever render markup or nothing -- an icon, or a conditional between
 * icons. Such a sibling does not put words beside the message, so it does not split it.
 *
 * This is the same judgement `wrap-jsx-text.ts` makes when it accepts a self-closing sibling
 * (`<Plus /> Create workspace`), widened by one shape that only appears around expressions: the
 * spinner-or-warning ternary that sits in front of a button label.
 */
function rendersOnlyMarkup(node: ts.Expression): boolean {
  const expression = unwrap(node);
  if (
    ts.isJsxElement(expression) || ts.isJsxSelfClosingElement(expression) ||
    ts.isJsxFragment(expression)
  ) {
    return true;
  }
  if (
    expression.kind === ts.SyntaxKind.NullKeyword ||
    expression.kind === ts.SyntaxKind.TrueKeyword ||
    expression.kind === ts.SyntaxKind.FalseKeyword ||
    (ts.isIdentifier(expression) && expression.text === "undefined")
  ) {
    return true;
  }
  if (ts.isConditionalExpression(expression)) {
    return rendersOnlyMarkup(expression.whenTrue) && rendersOnlyMarkup(expression.whenFalse);
  }
  if (ts.isBinaryExpression(expression)) {
    // `cond && <Icon />` renders the icon or nothing; `a || <Icon />` renders either side.
    if (expression.operatorToken.kind === ts.SyntaxKind.AmpersandAmpersandToken) {
      return rendersOnlyMarkup(expression.right);
    }
    if (
      expression.operatorToken.kind === ts.SyntaxKind.BarBarToken ||
      expression.operatorToken.kind === ts.SyntaxKind.QuestionQuestionToken
    ) {
      return rendersOnlyMarkup(expression.left) && rendersOnlyMarkup(expression.right);
    }
  }
  return false;
}

function isBlankText(node: ts.Node): boolean {
  return ts.isJsxText(node) && node.text.trim() === "";
}

/** Whether a sibling child of the message's container puts words beside it. */
function siblingRendersText(child: ts.JsxChild): boolean {
  if (isBlankText(child)) return false;
  if (ts.isJsxText(child)) return true;
  if (ts.isJsxSelfClosingElement(child)) return false;
  // `{/* comment */}` has no expression and renders nothing.
  if (ts.isJsxExpression(child)) return child.expression ? !rendersOnlyMarkup(child.expression) : false;
  return true;
}

/**
 * The whole value the literal is one alternative of.
 *
 * Climbs out of nested conditionals and `||`/`??`/`&&` defaults, because those only *choose* between
 * messages -- `{a ? 'x' : b ? 'y' : 'z'}` is one value with three branches. It stops at anything
 * else, so whatever sits above the result tells us what the literal is really being used for: a
 * `{...}` renders it, a template span embeds it in a longer sentence, a `+` glues it to something,
 * and a variable or a call takes it somewhere this transform cannot follow.
 */
function valueRoot(node: ts.Node): ts.Node {
  let current: ts.Node = node;
  for (let parent = current.parent; parent; parent = current.parent) {
    if (ts.isParenthesizedExpression(parent) || ts.isNonNullExpression(parent)) {
      current = parent;
      continue;
    }
    if (ts.isConditionalExpression(parent)) {
      // A literal in the condition is being tested, not shown.
      if (parent.condition === current) return current;
      current = parent;
      continue;
    }
    if (ts.isBinaryExpression(parent)) {
      const kind = parent.operatorToken.kind;
      // The left of `&&` is the guard; the right is the value.
      if (kind === ts.SyntaxKind.AmpersandAmpersandToken && parent.right === current) {
        current = parent;
        continue;
      }
      if (
        kind === ts.SyntaxKind.BarBarToken || kind === ts.SyntaxKind.QuestionQuestionToken
      ) {
        current = parent;
        continue;
      }
    }
    return current;
  }
  return current;
}

/** Whether every `${...}` in a template is a plain value reference: `${name}`, `${a.b?.c}`. */
function hasOnlySimpleSpans(node: ts.TemplateExpression): boolean {
  const isSimple = (expression: ts.Expression): boolean => {
    const inner = unwrap(expression);
    if (ts.isIdentifier(inner)) return true;
    if (ts.isPropertyAccessExpression(inner)) return isSimple(inner.expression);
    return false;
  };
  return node.templateSpans.every((span) => isSimple(span.expression));
}

/** Every condition a selector switches on, so they can be inspected for a count. */
function selectorConditions(root: ts.Expression): ts.Expression[] {
  const conditions: ts.Expression[] = [];
  const visit = (node: ts.Expression): void => {
    const expression = unwrap(node);
    if (ts.isConditionalExpression(expression)) {
      conditions.push(expression.condition);
      visit(expression.whenTrue);
      visit(expression.whenFalse);
      return;
    }
    if (ts.isBinaryExpression(expression)) {
      const kind = expression.operatorToken.kind;
      if (kind === ts.SyntaxKind.AmpersandAmpersandToken) {
        conditions.push(expression.left);
        visit(expression.right);
        return;
      }
      if (
        kind === ts.SyntaxKind.BarBarToken || kind === ts.SyntaxKind.QuestionQuestionToken
      ) {
        visit(expression.left);
        visit(expression.right);
      }
    }
  };
  visit(root);
  return conditions;
}

const COMPARISONS: readonly ts.SyntaxKind[] = [
  ts.SyntaxKind.EqualsEqualsToken,
  ts.SyntaxKind.EqualsEqualsEqualsToken,
  ts.SyntaxKind.ExclamationEqualsToken,
  ts.SyntaxKind.ExclamationEqualsEqualsToken,
  ts.SyntaxKind.GreaterThanToken,
  ts.SyntaxKind.GreaterThanEqualsToken,
  ts.SyntaxKind.LessThanToken,
  ts.SyntaxKind.LessThanEqualsToken,
];

/**
 * Whether a condition compares something against a number.
 *
 * `n === 1 ? 'person loses' : 'people lose'` is pluralization written as a ternary, and wrapping the
 * two branches separately ships a message pair that is only correct in languages with exactly two
 * plural forms -- while looking entirely correct in review. `kept === 0 ? … : …` is the same thing
 * with an ICU `=0` case. Both belong in a `plural` macro, which is a judgement about the sentence
 * rather than a span replacement.
 */
function comparesAgainstNumber(condition: ts.Expression): boolean {
  const expression = unwrap(condition);
  if (ts.isPrefixUnaryExpression(expression)) return comparesAgainstNumber(expression.operand);
  if (!ts.isBinaryExpression(expression)) return false;
  if (COMPARISONS.includes(expression.operatorToken.kind)) {
    return ts.isNumericLiteral(unwrap(expression.left)) ||
      ts.isNumericLiteral(unwrap(expression.right));
  }
  if (
    expression.operatorToken.kind === ts.SyntaxKind.AmpersandAmpersandToken ||
    expression.operatorToken.kind === ts.SyntaxKind.BarBarToken
  ) {
    return comparesAgainstNumber(expression.left) || comparesAgainstNumber(expression.right);
  }
  return false;
}

/** Every string literal in a value position of a selector, in source order. */
function selectorLiterals(root: ts.Expression): StringNode[] {
  const literals: StringNode[] = [];
  const visit = (node: ts.Expression): void => {
    const expression = unwrap(node);
    if (isStringNode(expression)) {
      literals.push(expression);
      return;
    }
    if (ts.isConditionalExpression(expression)) {
      visit(expression.whenTrue);
      visit(expression.whenFalse);
      return;
    }
    if (ts.isBinaryExpression(expression)) {
      const kind = expression.operatorToken.kind;
      if (kind === ts.SyntaxKind.AmpersandAmpersandToken) {
        visit(expression.right);
      } else if (
        kind === ts.SyntaxKind.BarBarToken || kind === ts.SyntaxKind.QuestionQuestionToken
      ) {
        visit(expression.left);
        visit(expression.right);
      }
    }
  };
  visit(root);
  return literals;
}

/** Whether the literal is built inside a function, so `t` runs at render rather than at import. */
function isInsideFunction(node: ts.Node): boolean {
  for (let current = node.parent; current; current = current.parent) {
    if (
      ts.isArrowFunction(current) || ts.isFunctionDeclaration(current) ||
      ts.isFunctionExpression(current) || ts.isMethodDeclaration(current) ||
      ts.isGetAccessor(current) || ts.isConstructorDeclaration(current)
    ) {
      return true;
    }
  }
  return false;
}

/** Where a refusal belongs in the follow-up worklist, and why. */
interface Refusal {
  reason: ResidueReason;
  /** The node whose whole rewrite this entry is part of. */
  groupAt: ts.Node;
}

/**
 * Why this literal cannot be wrapped on its own, or `undefined` if it can.
 *
 * `detected` is the set of literal offsets the detector reported, used to notice a selector that is
 * only partly covered.
 */
function refuse(
  node: StringNode, detected: ReadonlySet<number>, sourceFile: ts.SourceFile,
): Refusal | undefined {
  const root = valueRoot(node);
  const above = root.parent;

  // What the literal is used for, decided by what sits above the value it belongs to. Only a `{...}`
  // renders it as written; everything else composes it into something longer or carries it away.
  if (!above || !ts.isJsxExpression(above)) {
    if (above && (ts.isTemplateSpan(above) || ts.isTemplateExpression(above))) {
      // `` `Kept ${n} of ${total}` `` is one message; this literal is a piece of it.
      return {
        reason: "expression-fragment",
        groupAt: outermostOf(above, (n) => ts.isTemplateSpan(n) || ts.isTemplateExpression(n)),
      };
    }
    if (above && isConcatenation(above)) {
      // `'Removes it from ' + name + ' permanently.'` is one sentence in three pieces.
      return { reason: "expression-fragment", groupAt: outermostOf(above, isConcatenation) };
    }
    // Stored in a local, kept in a record, or handed to a call before it reaches the JSX. Whether
    // it stays a whole message depends on what happens to it next, which a span replacement cannot
    // see -- and a record's sibling properties are often discriminants (`value: 'disabled'`) that
    // must not move with it.
    return { reason: "expression-indirect-value", groupAt: above ?? root };
  }

  // A JSX element built at module scope is evaluated once, at import time, before a locale is
  // active. `t` there would freeze the message in whatever locale happened to load first.
  if (!isInsideFunction(node)) return { reason: "expression-module-scope", groupAt: above };

  // Grouped by the whole value rather than by the literal, so a refused branch and the branches
  // beside it are one task.
  if (ts.isTemplateExpression(node) && !hasOnlySimpleSpans(node)) {
    return { reason: "expression-complex-template", groupAt: root };
  }

  if (!ts.isTemplateExpression(node) && UNSAFE_IN_TEMPLATE.test(node.text)) {
    return { reason: "expression-unsafe-value", groupAt: root };
  }

  if (root !== node && ts.isExpression(root)) {
    if (selectorConditions(root).some(comparesAgainstNumber)) {
      return { reason: "expression-plural-selector", groupAt: root };
    }
    // Wrapping some branches of a selector and not others leaves the rest bare *and* invisible:
    // the guard reports what the detector finds, and it did not find these.
    const uncovered = selectorLiterals(root).some((literal) =>
      !detected.has(literal.getStart(sourceFile)) &&
      hasLetters(ts.isTemplateExpression(literal) ? literal.getText(sourceFile) : literal.text));
    if (uncovered) return { reason: "expression-partial-selector", groupAt: root };
  }

  const owner = above.parent;
  if (ts.isJsxElement(owner) || ts.isJsxFragment(owner)) {
    const splits = owner.children.some((child) => child !== above && siblingRendersText(child));
    if (splits) return { reason: "expression-fragment", groupAt: owner };
  }

  return undefined;
}

function isConcatenation(node: ts.Node): boolean {
  return ts.isBinaryExpression(node) && node.operatorToken.kind === ts.SyntaxKind.PlusToken;
}

/**
 * The outermost unbroken run of the same shape above `node`, so every piece of one message shares a
 * group and the follow-up pass rewrites them together.
 */
function outermostOf(node: ts.Node, matches: (node: ts.Node) => boolean): ts.Node {
  let outermost = node;
  for (let current = node.parent; current && matches(current); current = current.parent) {
    outermost = current;
  }
  return outermost;
}

function literalIndex(source: ts.SourceFile): Map<number, StringNode> {
  const index = new Map<number, StringNode>();
  const visit = (node: ts.Node): void => {
    if (isStringNode(node)) index.set(node.getStart(source), node);
    ts.forEachChild(node, visit);
  };
  visit(source);
  return index;
}

function residueOf(
  candidate: StringCandidate, refusal: Refusal, sourceFile: ts.SourceFile,
): ResidueEntry {
  return {
    kind: candidate.kind,
    attribute: candidate.attribute,
    parentTag: candidate.parentTag,
    text: candidate.text,
    line: candidate.line,
    column: candidate.column,
    reason: refusal.reason,
    group: `${candidate.file}:${refusal.groupAt.getStart(sourceFile)}`,
  };
}

/**
 * Plans a `` t`...` `` rewrite for every JSX-expression candidate that is a message on its own.
 *
 * `source` must be the text `candidates` were detected in, and `sourceFile` its parse. Candidates
 * that are not `jsx-expression` are ignored.
 */
export function planJsxExpressionWraps(
  source: string,
  sourceFile: ts.SourceFile,
  candidates: readonly StringCandidate[],
): Plan {
  const plan = emptyPlan();
  const index = literalIndex(sourceFile);
  const detected = new Set(
    candidates.filter((candidate) => candidate.kind === "jsx-expression")
      .map((candidate) => candidate.start),
  );

  const planned = candidates
    .filter((candidate) => candidate.kind === "jsx-expression")
    .map((candidate) => {
      const node = index.get(candidate.start);
      if (!node) {
        // A candidate with no matching node means the caller mismatched source and parse.
        throw new Error(`No string literal at offset ${candidate.start} in ${candidate.file}.`);
      }
      const root = valueRoot(node);
      return { candidate, node, root, refusal: refuse(node, detected, sourceFile) };
    });

  // A selector is one decision, so it is wrapped whole or not at all. Wrapping the branches that
  // pass while one is refused would leave a half-localized ternary in the source and a residue entry
  // that does not say its sibling already moved -- and the residue report is the only worklist the
  // follow-up pass gets.
  const brokenSelectors = new Set(
    planned.filter((entry) => entry.refusal && entry.root !== entry.node)
      .map((entry) => entry.root.getStart(sourceFile)),
  );

  for (const { candidate, node, root, refusal } of planned) {
    if (refusal) {
      plan.residue.push(residueOf(candidate, refusal, sourceFile));
      continue;
    }
    if (root !== node && brokenSelectors.has(root.getStart(sourceFile))) {
      plan.residue.push(
        residueOf(candidate, { reason: "expression-partial-selector", groupAt: root }, sourceFile),
      );
      continue;
    }
    plan.edits.push(wrapEdit(source, node, candidate));
  }

  return plan;
}

function wrapEdit(source: string, node: StringNode, candidate: StringCandidate): PlannedEdit {
  const before = source.slice(candidate.start, candidate.end);
  // A template already is a template: tagging it keeps every byte, including its substitutions and
  // any line breaks. A quoted string has to be retyped, and its cooked value is what it says --
  // `'It\'s ready'` becomes `` t`It's ready` ``, which is the same message.
  const replacement = ts.isStringLiteral(node)
    ? `${T_MACRO_SPECIFIER}\`${node.text}\``
    : `${T_MACRO_SPECIFIER}${before}`;

  return {
    kind: "jsx-expression",
    start: candidate.start,
    end: candidate.end,
    replacement,
    line: candidate.line,
    column: candidate.column,
    before,
  };
}

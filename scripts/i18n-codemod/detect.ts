// The one definition of "user-facing string candidate" in this repo. It has two consumers that
// must agree: `scripts/check-unlocalized-strings.ts` reports what is still bare, and the i18n
// codemod beside this file wraps what it safely can. If they each carried their own heuristics,
// the codemod's leftovers and the guard's findings would drift and neither number would mean
// anything.
//
// Rules only. Nothing here mutates source or knows about Lingui: a candidate is a span of text with
// a position, so a codemod can act on it with any AST library or with plain text edits.
//
// typescript6 = npm:typescript@6.0.3. TypeScript 7 (tsgo) ships no JS compiler API. This module
// parses and never type-checks, so a syntax-only `createSourceFile` is the whole dependency -- no
// program, no checker, no `tsconfig` resolution.
import { readdirSync, readFileSync } from "node:fs";
import { join, relative, resolve } from "node:path";
import ts from "typescript6";

/**
 * JSX attributes and component props whose string value is chrome the user reads.
 *
 * A whitelist, never a denylist: most attributes carry markup, styling or data, and the cost of
 * wrongly including one is a bad string in the catalog that a translator has to reject. Every entry
 * was checked against one question -- does this value render as text a user sees? -- by reading the
 * component that consumes it. Notable exclusions found by the same sweep: `className`, `rel`,
 * `sandbox`, and `d`, which holds SVG path data that reads like prose to a text heuristic but is
 * geometry.
 */
export const LOCALIZABLE_ATTRIBUTES: readonly string[] = [
  // Standard DOM attributes.
  "alt",
  "aria-description",
  "aria-label",
  "placeholder",
  "title",
  // Component props. `content` is Kumo's `<Tooltip>` text; the rest belong to this app's own
  // components, each of which renders the value directly -- `{label}`, `<h3>{heading}</h3>`,
  // `{actionLabel}`, `{isDeleting ? confirmingLabel : confirmLabel}`.
  "actionLabel",
  "ariaLabel",
  "confirmLabel",
  "confirmingLabel",
  "content",
  "description",
  "detail",
  "heading",
  "label",
  "message",
];

/** JSX elements that already localize every string node beneath them. */
export const LOCALIZED_CONTAINERS: readonly string[] = [
  "Plural",
  "Select",
  "SelectOrdinal",
  "Trans",
];

/**
 * HTML elements whose text is a technical literal -- a key name, a command, a log line -- rather
 * than prose. Translating what is inside them would break the thing they describe.
 */
export const TECHNICAL_LITERAL_ELEMENTS: readonly string[] = [
  "code",
  "kbd",
  "pre",
  "samp",
  "var",
];

/**
 * Brand and product names, which a translator can only ever copy through.
 *
 * Every entry is used in this repository: the third parties come from the `packages/gatekeeper-*`
 * integrations, the product names are ones the frontend renders, and the browsers and operating
 * systems are the values `errorReporting.ts` reports telemetry under. Generic words that merely
 * happen to name a package -- context, email, scheduler, portal -- are deliberately absent, because
 * they are ordinary copy wherever a user reads them.
 *
 * This only suppresses a string that is *nothing but* a brand. "Add credits in Cloudflare" is copy
 * and stays a candidate.
 */
export const PROPER_NOUNS: readonly string[] = [
  "AI Gateway",
  "Android",
  "Anthropic",
  "Chromium",
  "Claude",
  "Cloudflare",
  "Cloudflare Workers",
  "Confluence",
  "Firefox",
  "GitHub",
  "Google",
  "Google Calendar",
  "Google Docs",
  "Home Assistant",
  "Jira",
  "Linear",
  "Linux",
  "Microsoft",
  "Notion",
  "OpenAI",
  "Safari",
  "Slack",
  "Spotify",
  "Supabase",
  "Windows",
  "Workers",
  "Workers AI",
  "Wrangler",
  "ZoomInfo",
];

/** The repository root, derived from this file's own location. */
export const REPO_ROOT = resolve(import.meta.dirname, "..", "..");

/** The tree in scope for string extraction. */
export const SCAN_ROOT = join(REPO_ROOT, "packages", "workshop-frontend", "src");

/** Where intentional exceptions are recorded. */
export const DEFAULT_ALLOWLIST_PATH = join(import.meta.dirname, "unlocalized-allowlist.txt");

/**
 * Which syntax a candidate was found in.
 *
 * `string-literal` is reported for `.ts` files only. In a `.tsx` file the copy lives in JSX text
 * and in the whitelisted attributes, and the plain string literals around it are overwhelmingly
 * class names, ids, prop enums and config; scanning them there would bury the real findings and
 * would change the contract the codemod was verified against. A `.ts` file has no JSX to carry
 * copy, so a string literal is the only place copy can be.
 */
export type CandidateKind = "jsx-text" | "jsx-attribute" | "string-literal" | "jsx-expression";

/**
 * One string the detector considers user-facing, located precisely enough to rewrite.
 *
 * `start`/`end` bound the syntax node (a JSX text node, or a string literal including its quotes).
 * `textStart`/`textEnd` bound the text itself, which is the span a wrap replaces: for JSX text that
 * excludes the surrounding indentation, and for an attribute it excludes the quotes.
 */
export interface StringCandidate {
  kind: CandidateKind;
  /** Path as handed to the detector. {@link scanDirectory} makes it relative to its base. */
  file: string;
  /** The text with runs of whitespace collapsed to single spaces. What the rules judge. */
  text: string;
  /** The exact source between `textStart` and `textEnd`, whitespace intact. What a wrap preserves. */
  raw: string;
  /** Attribute name, for `jsx-attribute`. */
  attribute?: string;
  /** Tag name of the nearest enclosing element, for `jsx-text`. Empty inside a fragment. */
  parentTag?: string;
  /** Whether sibling children include a `{...}` expression, for `jsx-text`. */
  interpolated?: boolean;
  start: number;
  end: number;
  textStart: number;
  textEnd: number;
  /** 1-based, for human-readable reports. */
  line: number;
  /** 1-based. */
  column: number;
}

/** Intentional exceptions, parsed from an allowlist file. */
export interface Allowlist {
  /** Whole files, and directories when the entry ends in `/`. */
  files: Set<string>;
  /** Strings allowed in any file. */
  anywhere: Set<string>;
  /** Strings allowed in one named file. */
  perFile: Map<string, Set<string>>;
}

/** What a scan looks for, beyond the always-on rules. */
export interface DetectOptions {
  /**
   * Whether to report string literals inside JSX expression containers -- `{busy ? 'Saving…' :
   * 'Save'}` and `title={busy ? 'Saving…' : 'Save'}`.
   *
   * **On by default**, since `wrap-jsx-expressions.ts` transforms them. It was opt-in only while
   * the codemod had no transform for them, because a guard that reports what nothing can fix is a
   * guard people learn to ignore. Pass `false` to scan the way the guard did before that.
   */
  includeJsxExpressions?: boolean;

  /**
   * Whether to report string and template literals that a `.tsx` file builds *outside* JSX -- the
   * body of a helper that returns copy, such as `routes/outputs.tsx`'s `subtitle()`.
   *
   * Off by default, and sized before being trusted: the same rule over `.ts` files needed eleven
   * allowlist entries to reach a clean baseline, and a `.tsx` file holds far more non-copy literals
   * (props, ids, class names, query keys) than a `.ts` one. The rules applied are exactly the `.ts`
   * rules -- one heuristic, not a second copy of it.
   */
  includeHelperLiterals?: boolean;
}

/** Options for {@link scanDirectory}. */
export interface ScanOptions extends DetectOptions {
  /** Candidate paths are reported relative to this. Defaults to {@link REPO_ROOT}. */
  baseDir?: string;
  /** Exceptions to drop from the result. Defaults to none. */
  allowlist?: Allowlist;
}

const LETTER = /\p{L}/u;

// The punctuation UI copy trails a lone word with: "Loading…", "Authenticating...", "Ready?".
// No colon: "Note:" is indistinguishable from a URL scheme, and a lone word ending in one is a
// label for something that follows it, which belongs in the same message as whatever that is.
const TRAILING_UI_PUNCTUATION = /(?:\.{3}|[….!?])$/u;

// A single capitalized word, optionally carrying the intra-word punctuation English writes words
// with, and the punctuation UI copy trails it with: "Close", "Loading…", "Re-authenticating…",
// "Don't", "Sign-in". A hyphen or an apostrophe joins letters into one word rather than making two,
// so a rule that stopped at the first one was blind to a whole class of ordinary labels.
//
// Every segment after the first must be lowercase, which is what keeps identifiers out: "Content-Type"
// and "X-Frame-Options" capitalize each segment and are rejected, while "kebab-case" never starts
// with a capital in the first place.
const LONE_LABEL_WORD = /^\p{Lu}\p{Ll}*(?:['’-]\p{Ll}+)*(?:\.{3}|[….!?])?$/u;

// Four letters or more, counted across the whole word rather than the first segment, so "Add" and
// "New" stay out -- where a false positive is likelier than a real label -- while "Don't" and
// "Re-authenticating" are in.
const MINIMUM_LABEL_LETTERS = 4;
const LETTERS = /\p{L}/gu;

/** The shape half of the lone-word rule, shared so the two callers cannot drift apart. */
function looksLikeLabelWord(normalized: string): boolean {
  if (!LONE_LABEL_WORD.test(normalized)) return false;
  return (normalized.match(LETTERS)?.length ?? 0) >= MINIMUM_LABEL_LETTERS;
}

const PROPER_NOUN_SET = new Set(PROPER_NOUNS);

// A whole string that is a URL or a URL-ish scheme reference, with no room for prose around it.
const URL_LIKE = /^(?:[a-z][\d+.a-z-]*:(?:\/\/)?|\/\/)\S*$/i;

// A whole string that is a filesystem or route path: absolute, relative, or two or more
// slash-joined segments.
const PATH_LIKE = /^\.{0,2}\/\S*$|^[\w.@-]+(?:\/[\w.@-]+)+\/?$/;

// A bare filename with an extension: "logo.svg", "index.html".
const FILENAME_LIKE = /^[\w-]+\.[\da-z]{1,5}$/i;

const SNAKE_CASE = /^[\dA-Za-z]+(?:_[\dA-Za-z]+)+$/;
const KEBAB_CASE = /^[\da-z]+(?:-[\da-z]+)+$/;
// A dotted key such as `workshop.react-root`; segments may themselves be hyphenated or underscored.
const DOTTED = /^[\w-]+(?:\.[\w-]+)+$/;
const SCREAMING = /^[\dA-Z]+$/;
const WORD_CHARACTERS_ONLY = /^[A-Za-z][\dA-Za-z]*$/;
const UPPERCASE = /\p{Lu}/gu;
const LOWERCASE = /\p{Ll}/u;
const WHITESPACE_RUN = /\s+/gu;

/** Extensions {@link listSourceFiles} walks by default. */
export const SOURCE_FILE_EXTENSIONS: readonly string[] = [".ts", ".tsx"];

const EXCLUDED_FILE_PATTERNS: readonly RegExp[] = [
  /(?:^|\/)node_modules\//,
  /(?:^|\/)dist\//,
  // Tests hold fixture copy, not shipped copy.
  /\.(?:spec|test)\.[jt]sx?$/,
  // Generated: the router plugin rewrites its route tree on every build, so an edit would not last.
  /\.gen\.[jt]sx?$/,
  /(?:^|\/)generated\//,
  // Declarations hold types, never values, so nothing in one can reach a user.
  /\.d\.ts$/,
];

// Methods whose string argument is a needle matched against data rather than shown to anyone.
// `rpcErrors.ts` classifies backend failures by `.includes()` against runtime messages that read
// exactly like prose; those are the strings this keeps out.
const MATCH_METHODS: readonly string[] = [
  "endsWith",
  "includes",
  "indexOf",
  "lastIndexOf",
  "startsWith",
];

// Functions and tagged templates that already localize their argument: `t` and `msg` as tagged
// templates; `defineMessage` and the ICU-selector macros as calls. `plural`/`select`/`selectOrdinal`
// take their message text one level down, in an options object (`plural(n, { one: '...', other:
// '...' })`), not as a direct call argument -- see the object-literal-property branch below.
const LOCALIZING_CALLEES: readonly string[] = ["defineMessage", "msg", "plural", "select", "selectOrdinal", "t"];

// Operators whose string operand is a value being matched, never one being shown.
const EQUALITY_OPERATORS: readonly ts.SyntaxKind[] = [
  ts.SyntaxKind.EqualsEqualsToken,
  ts.SyntaxKind.EqualsEqualsEqualsToken,
  ts.SyntaxKind.ExclamationEqualsToken,
  ts.SyntaxKind.ExclamationEqualsEqualsToken,
];

/** Collapses runs of whitespace to single spaces and trims. The form every rule judges. */
export function normalizeText(text: string): string {
  return text.replaceAll(WHITESPACE_RUN, " ").trim();
}

/** Whether the string contains a letter in any script. Punctuation and digits alone do not count. */
export function hasLetters(text: string): boolean {
  return LETTER.test(text);
}

/** Whether the whole string is a URL rather than prose that mentions one. */
export function isUrlLike(text: string): boolean {
  return URL_LIKE.test(text);
}

/** Whether the whole string is a path or a bare filename. */
export function isPathLike(text: string): boolean {
  return PATH_LIKE.test(text) || FILENAME_LIKE.test(text);
}

/**
 * Whether the string is a programming identifier rather than copy: `camelCase`, `PascalCase`,
 * `snake_case`, `kebab-case`, `SCREAMING_CASE`, or a dotted key.
 *
 * Mixed case needs two capitals to count, so "Close" stays copy while "GitHub" and "onClick" do
 * not.
 */
export function isIdentifierLike(text: string): boolean {
  if (text.includes(" ")) return false;
  if (SCREAMING.test(text) || SNAKE_CASE.test(text) || KEBAB_CASE.test(text)) return true;
  if (DOTTED.test(text)) return true;
  if (!WORD_CHARACTERS_ONLY.test(text)) return false;
  const capitals = text.match(UPPERCASE)?.length ?? 0;
  return capitals >= 2 && LOWERCASE.test(text);
}

/**
 * Whether the string is nothing but a brand or product name. See {@link PROPER_NOUNS}.
 *
 * Trailing UI punctuation is ignored, so "Cloudflare…" is as much a brand as "Cloudflare".
 */
export function isProperNoun(text: string): boolean {
  return PROPER_NOUN_SET.has(normalizeText(text).replace(TRAILING_UI_PUNCTUATION, ""));
}

/**
 * Whether a lone word reads as a UI label: "Cancel", "Loading…", "Re-authenticating…", "Don't".
 *
 * Requires a capital and four letters in total, so "Add", "New" and "Esc" stay out, where a false
 * positive is likelier than a real label. Brands are excluded -- a translator can do nothing with
 * "Slack" but leave it, so putting it in the catalog only costs them a review.
 *
 * Callers must still exclude the element context this cannot see: inside `<code>` or `<kbd>` the
 * same word is a technical literal. {@link findStringCandidates} does that.
 */
export function isLoneLabelWord(text: string): boolean {
  const normalized = normalizeText(text);
  if (normalized.includes(" ")) return false;
  if (isUrlLike(normalized) || isPathLike(normalized) || isIdentifierLike(normalized)) return false;
  if (isProperNoun(normalized)) return false;
  return looksLikeLabelWord(normalized);
}

/**
 * Whether a run of text is copy a user reads: a phrase, or a lone word that reads as a label.
 *
 * The lone-word half is the part that needs guarding, and it is guarded by rule rather than by
 * word count: identifiers, paths, URLs and brands are already out, and the element context that
 * would make a word technical (`<code>`, `<kbd>`) is handled where the syntax tree is in scope.
 */
export function isUserFacingText(text: string): boolean {
  const normalized = normalizeText(text);
  if (!hasLetters(normalized)) return false;
  if (isUrlLike(normalized)) return false;
  if (isPathLike(normalized)) return false;
  if (isIdentifierLike(normalized)) return false;
  if (isProperNoun(normalized)) return false;
  return normalized.includes(" ") || looksLikeLabelWord(normalized);
}

/** Whether the attribute is one whose string value is read by a user. */
export function isLocalizableAttribute(name: string): boolean {
  return LOCALIZABLE_ATTRIBUTES.includes(name);
}

/** Whether an attribute's literal value is user-facing. */
export function isLocalizableAttributeValue(name: string, value: string): boolean {
  return isLocalizableAttribute(name) && isUserFacingText(value);
}

/** Whether this element localizes its own children, so the text inside it is already handled. */
export function isLocalizedContainer(tagName: string): boolean {
  return LOCALIZED_CONTAINERS.includes(tagName);
}

/** Whether this element's text is a technical literal rather than prose. */
export function isTechnicalLiteralElement(tagName: string): boolean {
  return TECHNICAL_LITERAL_ELEMENTS.includes(tagName);
}

/**
 * Whether the file is out of scope: dependencies, build output, tests, or generated code.
 * Takes a path in either separator style.
 */
export function isExcludedFile(filePath: string): boolean {
  const normalized = filePath.replaceAll("\\", "/");
  return EXCLUDED_FILE_PATTERNS.some((pattern) => pattern.test(normalized));
}

function tagNameOf(node: ts.JsxOpeningElement | ts.JsxClosingElement, source: ts.SourceFile): string {
  return node.tagName.getText(source);
}

function hasLocalizedAncestor(node: ts.Node, source: ts.SourceFile): boolean {
  for (let current = node.parent; current; current = current.parent) {
    if (ts.isJsxElement(current) && isLocalizedContainer(tagNameOf(current.openingElement, source))) {
      return true;
    }
  }
  return false;
}

// Suppresses the whole text, not just a lone word in it: a sentence inside `<pre>` is a log line,
// and a sentence inside `<code>` is a snippet. Attributes are deliberately not suppressed -- an
// `aria-label` on a `<code>` element is still copy.
function hasTechnicalLiteralAncestor(node: ts.Node, source: ts.SourceFile): boolean {
  for (let current = node.parent; current; current = current.parent) {
    if (
      ts.isJsxElement(current) &&
      isTechnicalLiteralElement(tagNameOf(current.openingElement, source))
    ) {
      return true;
    }
  }
  return false;
}

function hasExpressionSibling(node: ts.Node): boolean {
  const parent = node.parent;
  if (!parent || !(ts.isJsxElement(parent) || ts.isJsxFragment(parent))) return false;
  return parent.children.some((child) => ts.isJsxExpression(child) && child.expression !== undefined);
}

function enclosingTag(node: ts.Node, source: ts.SourceFile): string | undefined {
  const parent = node.parent;
  if (!parent) return undefined;
  if (ts.isJsxElement(parent)) return tagNameOf(parent.openingElement, source);
  if (ts.isJsxFragment(parent)) return "";
  return undefined;
}

function stringLiteralOf(initializer: ts.JsxAttribute["initializer"]): ts.StringLiteral | undefined {
  if (!initializer) return undefined;
  if (ts.isStringLiteral(initializer)) return initializer;
  if (ts.isJsxExpression(initializer) && initializer.expression &&
      ts.isStringLiteral(initializer.expression)) {
    return initializer.expression;
  }
  return undefined;
}

function positionOf(source: ts.SourceFile, offset: number): { line: number; column: number } {
  const { line, character } = source.getLineAndCharacterOfPosition(offset);
  return { line: line + 1, column: character + 1 };
}

function collectJsxText(
  node: ts.JsxText, file: string, sourceText: string, source: ts.SourceFile,
): StringCandidate | undefined {
  // `node.pos`/`node.end` rather than `getStart()`: in JSX the surrounding whitespace is part of
  // the text node, not trivia to skip past, and a wrap has to know where it really begins.
  const rawNode = sourceText.slice(node.pos, node.end);
  const leading = rawNode.length - rawNode.trimStart().length;
  const trailing = rawNode.length - rawNode.trimEnd().length;
  if (leading + trailing >= rawNode.length) return undefined;

  const textStart = node.pos + leading;
  const textEnd = node.end - trailing;
  const raw = sourceText.slice(textStart, textEnd);
  const text = normalizeText(raw);
  if (!isUserFacingText(text)) return undefined;
  if (hasLocalizedAncestor(node, source)) return undefined;
  if (hasTechnicalLiteralAncestor(node, source)) return undefined;

  return {
    kind: "jsx-text",
    file,
    text,
    raw,
    parentTag: enclosingTag(node, source),
    interpolated: hasExpressionSibling(node),
    start: node.pos,
    end: node.end,
    textStart,
    textEnd,
    ...positionOf(source, textStart),
  };
}

function collectJsxAttribute(
  node: ts.JsxAttribute, file: string, sourceText: string, source: ts.SourceFile,
): StringCandidate | undefined {
  const name = node.name.getText(source);
  const literal = stringLiteralOf(node.initializer);
  if (!literal) return undefined;
  if (!isLocalizableAttributeValue(name, literal.text)) return undefined;

  const start = literal.getStart(source);
  const end = literal.getEnd();
  // Inside the quotes: a wrap replaces the value, never the delimiters.
  const textStart = start + 1;
  const textEnd = end - 1;

  return {
    kind: "jsx-attribute",
    file,
    text: normalizeText(literal.text),
    raw: sourceText.slice(textStart, textEnd),
    attribute: name,
    start,
    end,
    textStart,
    textEnd,
    ...positionOf(source, start),
  };
}

type StringNode = ts.StringLiteral | ts.NoSubstitutionTemplateLiteral | ts.TemplateExpression;

function isStringNode(node: ts.Node): node is StringNode {
  return ts.isStringLiteral(node) || ts.isNoSubstitutionTemplateLiteral(node) ||
    ts.isTemplateExpression(node);
}

/**
 * What a string node says, with each `${...}` standing in as `{0}`.
 *
 * The placeholder deliberately carries no space. A URL built by template --
 * `` `https://dash.cloudflare.com/?to=/${account}/ai` `` -- has to keep reading as one unbroken
 * token so {@link isUrlLike} still recognises it; joining the pieces with a space would turn every
 * such template into something that looks like a sentence.
 */
function textOfStringNode(node: StringNode): string {
  if (!ts.isTemplateExpression(node)) return node.text;
  return node.head.text + node.templateSpans.map((span) => `{0}${span.literal.text}`).join("");
}

function calleeName(node: ts.Node): string | undefined {
  if (ts.isIdentifier(node)) return node.text;
  if (ts.isPropertyAccessExpression(node)) return node.name.text;
  return undefined;
}

/**
 * Why a string literal in a `.ts` file is not copy, structurally.
 *
 * Each rule answers "could this value ever reach a user?" from syntax alone, which is all a
 * parse-only detector has. Anything a rule cannot settle stays a finding -- a false positive is
 * visible and gets an allowlist line, whereas an over-broad rule hides real copy silently.
 */
function isNonCopyPosition(node: StringNode): boolean {
  const parent = node.parent;
  if (!parent) return false;

  // A type, not a value: `type DiffStatus = 'Added' | 'Deleted'`.
  if (ts.isLiteralTypeNode(parent)) return true;
  // A module specifier: `import x from './x'`, `export * from './x'`, `import('./x')`.
  if (ts.isImportDeclaration(parent) || ts.isExportDeclaration(parent)) return true;
  // The name side of a property, not its value: `{ 'a phrase': 1 }` and `obj['a phrase']`.
  if ((ts.isPropertyAssignment(parent) || ts.isPropertySignature(parent)) && parent.name === node) {
    return true;
  }
  if (ts.isElementAccessExpression(parent) && parent.argumentExpression === node) return true;
  // An operand of an equality test: `model.status !== 'Modified'`. The literal is being compared
  // against a value, so translating it would break the comparison and would never be read by
  // anyone. Same class as the `.includes()` rule below, reached through an operator instead of a
  // method.
  if (
    ts.isBinaryExpression(parent) && EQUALITY_OPERATORS.includes(parent.operatorToken.kind) &&
    (parent.left === node || parent.right === node)
  ) {
    return true;
  }
  // A `switch` label, which is the same comparison written another way.
  if (ts.isCaseClause(parent) && parent.expression === node) return true;
  // A CSS class list: `el.className = 'row omitted-row'`.
  if (
    ts.isBinaryExpression(parent) && parent.right === node &&
    parent.operatorToken.kind === ts.SyntaxKind.EqualsToken &&
    ts.isPropertyAccessExpression(parent.left) && parent.left.name.text === "className"
  ) {
    return true;
  }

  if (ts.isCallExpression(parent) && parent.arguments.includes(node as ts.Expression)) {
    // Log text is written for whoever reads the console, not for a user.
    if (ts.isPropertyAccessExpression(parent.expression) &&
      ts.isIdentifier(parent.expression.expression) &&
      parent.expression.expression.text === "console") {
      return true;
    }
    const callee = calleeName(parent.expression);
    // A needle compared against data, never rendered.
    if (callee && MATCH_METHODS.includes(callee)) return true;
    // Already localized.
    if (callee && LOCALIZING_CALLEES.includes(callee)) return true;
    if (callee === "_" && ts.isPropertyAccessExpression(parent.expression) &&
      ts.isIdentifier(parent.expression.expression) &&
      parent.expression.expression.text === "i18n") {
      return true;
    }
  }
  // Already localized: t`...`, msg`...`.
  if (ts.isTaggedTemplateExpression(parent) && parent.template === node) {
    const tag = calleeName(parent.tag);
    if (tag && LOCALIZING_CALLEES.includes(tag)) return true;
  }
  // Already localized: a message inside the options object of `plural(value, { one: '...', other:
  // '...' })` / `select(...)` / `selectOrdinal(...)`. The string is a property value one level
  // below the call (inside its options object), not a direct call argument, so the call-argument
  // check above cannot see it.
  if (
    ts.isPropertyAssignment(parent) && parent.initializer === node &&
    ts.isObjectLiteralExpression(parent.parent) &&
    ts.isCallExpression(parent.parent.parent) &&
    parent.parent.parent.arguments.includes(parent.parent as ts.Expression)
  ) {
    const callee = calleeName(parent.parent.parent.expression);
    if (callee && LOCALIZING_CALLEES.includes(callee)) return true;
  }
  // A dynamic import specifier.
  if (ts.isCallExpression(parent) && parent.expression.kind === ts.SyntaxKind.ImportKeyword) {
    return true;
  }
  return false;
}

function collectStringLiteral(
  node: StringNode, file: string, sourceText: string, source: ts.SourceFile,
): StringCandidate | undefined {
  const text = textOfStringNode(node);
  if (!isUserFacingText(text)) return undefined;
  if (isNonCopyPosition(node)) return undefined;

  const start = node.getStart(source);
  const end = node.getEnd();
  // A plain literal's delimiters are not part of the message; a template's placeholders are woven
  // through it, so there is no inner span to point at and the whole node is reported.
  const contiguous = !ts.isTemplateExpression(node);
  const textStart = contiguous ? start + 1 : start;
  const textEnd = contiguous ? end - 1 : end;

  return {
    kind: "string-literal",
    file,
    text: normalizeText(text),
    raw: sourceText.slice(textStart, textEnd),
    start,
    end,
    textStart,
    textEnd,
    ...positionOf(source, start),
  };
}

/**
 * The JSX expression container a node sits inside, and what that container belongs to.
 *
 * A container that is an element's child renders whatever it evaluates to, so a string literal in
 * one is copy. A container that is an attribute value is only copy when the attribute is one that
 * carries copy -- `className={active ? 'row on' : 'row'}` is styling, and there are far more of
 * those than of the real thing.
 */
function jsxExpressionContext(
  node: ts.Node, source: ts.SourceFile,
): { attribute?: string; parentTag?: string } | undefined {
  for (let current = node.parent; current; current = current.parent) {
    // Stop at the first construct that is not part of one expression: a nested element's children
    // and a nested arrow's body are their own contexts, handled when the walk reaches them.
    if (ts.isJsxElement(current) || ts.isJsxFragment(current) || ts.isJsxSelfClosingElement(current)) {
      return undefined;
    }
    if (!ts.isJsxExpression(current)) continue;

    const owner = current.parent;
    if (ts.isJsxAttribute(owner)) {
      const name = owner.name.getText(source);
      return isLocalizableAttribute(name) ? { attribute: name } : undefined;
    }
    if (ts.isJsxElement(owner)) return { parentTag: tagNameOf(owner.openingElement, source) };
    if (ts.isJsxFragment(owner)) return { parentTag: "" };
    return undefined;
  }
  return undefined;
}

/** Whether the node sits anywhere inside JSX, as opposed to in ordinary module code beside it. */
function isInsideJsx(node: ts.Node): boolean {
  for (let current = node.parent; current; current = current.parent) {
    if (
      ts.isJsxExpression(current) || ts.isJsxAttribute(current) || ts.isJsxElement(current) ||
      ts.isJsxSelfClosingElement(current) || ts.isJsxFragment(current)
    ) {
      return true;
    }
  }
  return false;
}

function collectJsxExpressionLiteral(
  node: StringNode, file: string, sourceText: string, source: ts.SourceFile,
): StringCandidate | undefined {
  // `title={"Close"}` is already a `jsx-attribute` candidate; reporting it again under a second
  // kind would double-count it and give the codemod two plans for one span.
  //
  // Only a quoted string is, though. `stringLiteralOf` accepts nothing else, so a template value --
  // `` title={`Actions for ${name}`} `` -- reaches no other collector, and excluding it here on the
  // strength of its position alone made it invisible to both the codemod and the guard. Templates
  // therefore fall through and are wrapped as expressions, which is also the shape that fits: `t`
  // tags the existing template and keeps every byte, where the attribute path retypes the value.
  if (
    ts.isStringLiteral(node) &&
    ts.isJsxExpression(node.parent) && node.parent.expression === node &&
    ts.isJsxAttribute(node.parent.parent)
  ) {
    return undefined;
  }

  const context = jsxExpressionContext(node, source);
  if (!context) return undefined;

  const text = textOfStringNode(node);
  if (!isUserFacingText(text)) return undefined;
  if (isNonCopyPosition(node)) return undefined;
  if (hasLocalizedAncestor(node, source)) return undefined;
  // An attribute keeps its meaning inside `<code>`; rendered children do not.
  if (context.attribute === undefined && hasTechnicalLiteralAncestor(node, source)) return undefined;

  const start = node.getStart(source);
  const end = node.getEnd();
  const contiguous = !ts.isTemplateExpression(node);
  const textStart = contiguous ? start + 1 : start;
  const textEnd = contiguous ? end - 1 : end;

  return {
    kind: "jsx-expression",
    file,
    text: normalizeText(text),
    raw: sourceText.slice(textStart, textEnd),
    ...context,
    start,
    end,
    textStart,
    textEnd,
    ...positionOf(source, start),
  };
}

/**
 * Every user-facing string in one file, in source order.
 *
 * The rules applied depend on the extension. A `.tsx` file is scanned for JSX text and whitelisted
 * attributes; a `.ts` file is scanned for string and template literals, since it has no JSX to
 * carry copy. `file` decides which, and is also the label on the results, so a caller can pass
 * whatever path form its report should show. It does NOT apply {@link isExcludedFile} or an
 * allowlist; {@link scanDirectory} composes those.
 */
export function findStringCandidates(
  file: string, sourceText: string, options: DetectOptions = {},
): StringCandidate[] {
  const jsx = file.endsWith(".tsx");
  const source = ts.createSourceFile(
    file, sourceText, ts.ScriptTarget.Latest, /* setParentNodes */ true,
    jsx ? ts.ScriptKind.TSX : ts.ScriptKind.TS,
  );
  const candidates: StringCandidate[] = [];

  const visit = (node: ts.Node): void => {
    if (!jsx) {
      if (isStringNode(node)) {
        const candidate = collectStringLiteral(node, file, sourceText, source);
        if (candidate) candidates.push(candidate);
        // A template expression's own head and spans are not separate candidates.
        if (ts.isTemplateExpression(node)) {
          for (const span of node.templateSpans) visit(span.expression);
          return;
        }
      }
    } else if (ts.isJsxText(node)) {
      const candidate = collectJsxText(node, file, sourceText, source);
      if (candidate) candidates.push(candidate);
    } else if (ts.isJsxAttribute(node)) {
      const candidate = collectJsxAttribute(node, file, sourceText, source);
      if (candidate) candidates.push(candidate);
    } else if (isStringNode(node)) {
      // Inside JSX the expression rule owns the literal; outside it, the file is behaving like a
      // `.ts` module and the `.ts` rule is the right one. The two are disjoint by construction.
      const candidate = isInsideJsx(node)
        ? (options.includeJsxExpressions === false
          ? undefined
          : collectJsxExpressionLiteral(node, file, sourceText, source))
        : (options.includeHelperLiterals
          ? collectStringLiteral(node, file, sourceText, source)
          : undefined);
      if (candidate) candidates.push(candidate);
      if (ts.isTemplateExpression(node)) {
        for (const span of node.templateSpans) visit(span.expression);
        return;
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(source);

  return candidates;
}

/** An allowlist matching nothing. */
export function emptyAllowlist(): Allowlist {
  return { files: new Set(), anywhere: new Set(), perFile: new Map() };
}

/**
 * Parses an allowlist file.
 *
 * One entry per line; `#` starts a comment and blank lines are ignored. An entry is a path, or a
 * path and a string separated by `|`:
 *
 * ```text
 * packages/workshop-frontend/src/legacy/      # every file under a directory
 * packages/workshop-frontend/src/Debug.tsx    # one file
 * packages/workshop-frontend/src/Api.tsx|GET /v1/things   # one string in one file
 * *|Cloudflare Workers                        # one string anywhere
 * ```
 *
 * The string is matched against a candidate's normalized `text`, so leading, trailing, and repeated
 * whitespace in the source does not have to be reproduced here.
 */
export function parseAllowlist(contents: string): Allowlist {
  const allowlist = emptyAllowlist();

  for (const rawLine of contents.split("\n")) {
    const line = rawLine.replace(/(^|\s)#.*$/, "").trim();
    if (!line) continue;

    const separator = line.indexOf("|");
    if (separator === -1) {
      allowlist.files.add(line.replaceAll("\\", "/"));
      continue;
    }

    const file = line.slice(0, separator).trim().replaceAll("\\", "/");
    const text = normalizeText(line.slice(separator + 1));
    if (!text) continue;

    if (file === "*") {
      allowlist.anywhere.add(text);
      continue;
    }
    const existing = allowlist.perFile.get(file);
    if (existing) existing.add(text);
    else allowlist.perFile.set(file, new Set([text]));
  }

  return allowlist;
}

/** Reads an allowlist file. A missing file is an empty allowlist, not an error. */
export function loadAllowlist(path: string = DEFAULT_ALLOWLIST_PATH): Allowlist {
  try {
    return parseAllowlist(readFileSync(path, "utf8"));
  } catch (error) {
    if ((error as NodeJS.ErrnoException | undefined)?.code === "ENOENT") return emptyAllowlist();
    throw error;
  }
}

/** Whether an allowlist covers this candidate. */
export function isAllowlisted(allowlist: Allowlist, candidate: StringCandidate): boolean {
  const file = candidate.file.replaceAll("\\", "/");
  if (allowlist.files.has(file)) return true;
  for (const entry of allowlist.files) {
    if (entry.endsWith("/") && file.startsWith(entry)) return true;
  }
  if (allowlist.anywhere.has(candidate.text)) return true;
  return allowlist.perFile.get(file)?.has(candidate.text) ?? false;
}

/**
 * Every in-scope source file under `rootDir`, absolute and sorted for a stable report order.
 *
 * `extensions` narrows the walk. The codemod passes `[".tsx"]` because its transforms are JSX
 * rewrites and it has nothing to do with a `.ts` file; the guard takes the default and reports on
 * both.
 */
export function listSourceFiles(
  rootDir: string, extensions: readonly string[] = SOURCE_FILE_EXTENSIONS,
): string[] {
  const found: string[] = [];

  const walk = (directory: string): void => {
    // Sorted by code unit rather than `localeCompare`, so the report order is the same everywhere.
    const entries = readdirSync(directory, { withFileTypes: true }).toSorted((a, b) =>
      a.name < b.name ? -1 : a.name > b.name ? 1 : 0);
    for (const entry of entries) {
      const path = join(directory, entry.name);
      if (entry.isDirectory()) {
        if (!isExcludedFile(`${path}/`)) walk(path);
      } else if (extensions.some((ext) => entry.name.endsWith(ext)) && !isExcludedFile(path)) {
        found.push(path);
      }
    }
  };
  walk(rootDir);

  return found;
}

/**
 * Every user-facing string still bare under `rootDir`.
 *
 * The composed pipeline both consumers run, so that "what the codemod left behind" and "what the
 * guard reports" are the same question asked twice.
 */
export function scanDirectory(rootDir: string, options: ScanOptions = {}): StringCandidate[] {
  const baseDir = options.baseDir ?? REPO_ROOT;
  const allowlist = options.allowlist;
  const candidates: StringCandidate[] = [];

  for (const path of listSourceFiles(rootDir)) {
    const file = relative(baseDir, path).replaceAll("\\", "/");
    const found = findStringCandidates(file, readFileSync(path, "utf8"), options);
    for (const candidate of found) {
      if (allowlist && isAllowlisted(allowlist, candidate)) continue;
      candidates.push(candidate);
    }
  }

  return candidates;
}

// The vocabulary the three planners share: a span replacement, and a reason for not making one.
//
// The codemod rewrites source by splicing text at offsets rather than by printing a syntax tree.
// `detect.ts` already reports exact offsets, every transform here is a local span replacement, and
// a printer would be free to reformat -- which is the one thing a JSX wrap must never do, because
// whitespace inside JSX children is rendered. Applying edits from the end of the file backwards
// keeps every offset valid without bookkeeping.
import type { CandidateKind } from "./detect.ts";

/** A replacement of `[start, end)` with `replacement`. An insertion has `start === end`. */
export interface Edit {
  start: number;
  end: number;
  replacement: string;
}

/** Which transform produced an edit. */
export type EditKind = "jsx-text" | "jsx-attribute" | "jsx-expression" | "import";

/** An edit with everything a dry-run report needs to be reviewed by a human. */
export interface PlannedEdit extends Edit {
  kind: EditKind;
  /** 1-based. */
  line: number;
  /** 1-based. */
  column: number;
  /** The source the edit replaces. */
  before: string;
}

/**
 * Why a detected string was left alone. Each value is a distinct unit of work for the follow-up
 * pass that handles what the codemod will not.
 */
export type ResidueReason =
  /** A sibling `{expression}`: the message spans the interpolation and must be rewritten whole. */
  | "interpolated"
  /** A sibling element with its own children: the sentence is split across markup. */
  | "split-by-markup"
  /** JSX text with syntax a `<Trans>` wrap would change the meaning of. */
  | "unsupported-text-content"
  /** An attribute value that would not survive being retyped as a template literal. */
  | "attribute-unsafe-value"
  /**
   * A literal inside a JSX expression that is part of a longer message: a piece of a template, an
   * operand of a concatenation, or a branch whose sentence continues in the sibling JSX.
   */
  | "expression-fragment"
  /**
   * A ternary that selects copy on a numeric comparison. It is pluralization written as a
   * conditional and belongs in a `plural` macro; wrapping the branches separately is only correct
   * in languages with exactly two plural forms.
   */
  | "expression-plural-selector"
  /**
   * A selector whose other branches hold text the detector did not report, so wrapping the branches
   * it did report would leave the rest bare and invisible to the guard.
   */
  | "expression-partial-selector"
  /**
   * A literal that reaches the JSX indirectly -- through a local, a record of options, or a call --
   * so whether it stays a whole message depends on what happens to it next.
   */
  | "expression-indirect-value"
  /** A template whose substitutions are more than plain value references. */
  | "expression-complex-template"
  /** A JSX expression evaluated at module scope, where no locale is active yet. */
  | "expression-module-scope"
  /** An expression value that would not survive being retyped as a template literal. */
  | "expression-unsafe-value"
  /** The file already binds `t` or `Trans` to something else. */
  | "identifier-conflict"
  /** The rewritten file did not parse, so the whole file was left untouched. */
  | "output-parse-failed";

/** One detected string the codemod declined to transform. */
export interface ResidueEntry {
  kind: CandidateKind;
  /** For `jsx-attribute`. */
  attribute?: string;
  /** For `jsx-text`. */
  parentTag?: string;
  text: string;
  /** 1-based. */
  line: number;
  /** 1-based. */
  column: number;
  reason: ResidueReason;
  /**
   * Entries that must be rewritten together share this, because they are fragments of one message.
   * `"<file>:<offset of the enclosing element>"`.
   */
  group?: string;
}

/** What a planner returns: the edits it is confident in, and what it refused and why. */
export interface Plan {
  edits: PlannedEdit[];
  residue: ResidueEntry[];
}

/** An empty plan, so a caller can merge unconditionally. */
export function emptyPlan(): Plan {
  return { edits: [], residue: [] };
}

/**
 * Applies edits to `source`.
 *
 * Edits are applied from the highest offset down, so each one lands on offsets that the earlier
 * ones have not moved. Overlapping edits are a bug in a planner rather than a case to resolve, so
 * they throw.
 */
export function applyEdits(source: string, edits: readonly Edit[]): string {
  const ordered = [...edits].toSorted((a, b) => b.start - a.start || b.end - a.end);

  let result = source;
  let lowestApplied = Number.POSITIVE_INFINITY;
  for (const edit of ordered) {
    if (edit.start > edit.end) {
      throw new Error(`Inverted edit at ${edit.start}-${edit.end}.`);
    }
    if (edit.end > lowestApplied) {
      throw new Error(`Overlapping edits at offset ${edit.start}-${edit.end}.`);
    }
    result = result.slice(0, edit.start) + edit.replacement + result.slice(edit.end);
    lowestApplied = edit.start;
  }

  return result;
}

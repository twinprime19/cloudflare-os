import assert from "node:assert/strict";
import { describe, it } from "node:test";
import ts from "typescript6";
import { findStringCandidates } from "./detect.ts";
import { applyEdits } from "./edits.ts";
import { planJsxTextWraps } from "./wrap-jsx-text.ts";

const FILE = "packages/workshop-frontend/src/Sample.tsx";

function plan(source: string) {
  const sourceFile = ts.createSourceFile(FILE, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
  return planJsxTextWraps(source, sourceFile, findStringCandidates(FILE, source));
}

function wrapped(source: string): string {
  return applyEdits(source, plan(source).edits);
}

function reasons(source: string): string[] {
  return plan(source).residue.map((entry) => entry.reason);
}

describe("planJsxTextWraps", () => {
  it("wraps a phrase that is the whole of its element", () => {
    assert.equal(wrapped("<p>Delete this workspace?</p>"),
      "<p><Trans>Delete this workspace?</Trans></p>");
  });

  it("wraps a lone label word", () => {
    assert.equal(wrapped("<button>Cancel</button>"), "<button><Trans>Cancel</Trans></button>");
  });

  it("leaves the indentation outside the wrap and the text inside it byte-identical", () => {
    const source = "<p>\n      Delete this\n      workspace?\n    </p>";
    assert.equal(wrapped(source),
      "<p>\n      <Trans>Delete this\n      workspace?</Trans>\n    </p>");
  });

  it("preserves the spaces JSX renders around text on one line", () => {
    // A space beside a tag on the same line is rendered, so it must stay outside the wrap.
    assert.equal(wrapped("<span> Saved </span>"), "<span> <Trans>Saved</Trans> </span>");
  });

  it("changes nothing but the wrap, byte for byte", () => {
    const source = "const a = <p className=\"x\">Ready to go</p>\r\n";
    const output = wrapped(source);
    assert.equal(output.replace("<Trans>", "").replace("</Trans>", ""), source);
  });

  it("reports text broken by an interpolation instead of wrapping the pieces", () => {
    const source = "<p>Always approve {label} for this gadget?</p>";
    const result = plan(source);
    assert.deepEqual(result.edits, []);
    assert.deepEqual(result.residue.map((entry) => entry.reason), ["interpolated", "interpolated"]);
    // Both fragments belong to one message, so they are one unit of work.
    assert.equal(new Set(result.residue.map((entry) => entry.group)).size, 1);
  });

  it("reports a sentence split by an element that renders text of its own", () => {
    assert.deepEqual(reasons("<p>Read the <a href=\"/docs\">documentation</a></p>"),
      ["split-by-markup"]);
  });

  it("reports both runs when a self-closing element sits inside a sentence", () => {
    assert.deepEqual(reasons("<p>See below <br /> for the details</p>"),
      ["split-by-markup", "split-by-markup"]);
  });

  it("wraps a label whose only sibling is a self-closing element", () => {
    // An icon renders no words, so the label is still the whole message.
    assert.equal(wrapped("<button><Plus size={14} /> Create workspace</button>"),
      "<button><Plus size={14} /> <Trans>Create workspace</Trans></button>");
  });

  it("reports text carrying syntax a Trans wrap would reinterpret", () => {
    assert.deepEqual(reasons("<p>Ends with a brace } and more</p>"), ["unsupported-text-content"]);
  });

  it("ignores attribute candidates", () => {
    assert.deepEqual(plan('<input placeholder="Ask anything" />').edits, []);
  });

  it("refuses candidates that do not belong to the source it was given", () => {
    const source = "<p>Delete this workspace?</p>";
    const sourceFile = ts.createSourceFile(FILE, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
    const foreign = findStringCandidates(FILE, "<div>\n  <p>A completely different tree</p>\n</div>");
    assert.throws(() => planJsxTextWraps(source, sourceFile, foreign), /No JSX text node/);
  });

  it("leaves nothing to do on its own output", () => {
    const source = "<p>Delete this workspace?</p>\n<button>Cancel</button>";
    const once = wrapped(source);
    assert.deepEqual(plan(once).edits, []);
    assert.equal(wrapped(once), once);
  });
});

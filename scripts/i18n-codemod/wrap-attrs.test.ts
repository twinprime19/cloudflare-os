import assert from "node:assert/strict";
import { describe, it } from "node:test";
import ts from "typescript6";
import { findStringCandidates } from "./detect.ts";
import { applyEdits } from "./edits.ts";
import { planAttributeWraps } from "./wrap-attrs.ts";

const FILE = "packages/workshop-frontend/src/Sample.tsx";

function plan(source: string) {
  const sourceFile = ts.createSourceFile(FILE, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
  return planAttributeWraps(source, sourceFile, findStringCandidates(FILE, source));
}

function wrapped(source: string): string {
  return applyEdits(source, plan(source).edits);
}

describe("planAttributeWraps", () => {
  it("gives a bare string value the braces a macro call needs", () => {
    assert.equal(wrapped('<input placeholder="Ask anything" />'),
      "<input placeholder={t`Ask anything`} />");
  });

  it("reuses the braces a value already has", () => {
    assert.equal(wrapped('<button aria-label={"Close"} />'), "<button aria-label={t`Close`} />");
  });

  it("keeps the value byte-identical, including a quote of the other kind", () => {
    assert.equal(wrapped(`<a title='Say "hello"' />`), "<a title={t`Say \"hello\"`} />");
  });

  it("wraps every whitelisted attribute on one element", () => {
    assert.equal(
      wrapped('<img alt="A logo" title="Our logo" />'),
      "<img alt={t`A logo`} title={t`Our logo`} />",
    );
  });

  it("leaves values that a template literal would read differently", () => {
    for (const source of [
      '<a title="Costs 50\\u0025 less" />',
      "<a title=\"Tom &amp; Jerry drink tea\" />",
      '<a title="Use a `backtick` here" />',
      '<a title="Braces {like} these" />',
    ]) {
      const result = plan(source);
      assert.deepEqual(result.edits, [], source);
      assert.deepEqual(result.residue.map((entry) => entry.reason), ["attribute-unsafe-value"],
        source);
    }
  });

  it("records which attribute it declined", () => {
    const [residue] = plan('<a title="Braces {like} these" />').residue;
    assert.equal(residue.attribute, "title");
    assert.equal(residue.kind, "jsx-attribute");
  });

  it("ignores JSX text candidates", () => {
    assert.deepEqual(plan("<p>Delete this workspace?</p>").edits, []);
  });

  it("leaves nothing to do on its own output", () => {
    const once = wrapped('<input placeholder="Ask anything" aria-label="Search" />');
    assert.deepEqual(plan(once).edits, []);
    assert.equal(wrapped(once), once);
  });
});

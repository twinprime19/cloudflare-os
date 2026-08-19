import assert from "node:assert/strict";
import { describe, it } from "node:test";
import ts from "typescript6";
import { findStringCandidates } from "./detect.ts";
import { applyEdits } from "./edits.ts";
import { planJsxExpressionWraps } from "./wrap-jsx-expressions.ts";

const FILE = "packages/workshop-frontend/src/Sample.tsx";

function plan(source: string) {
  const sourceFile = ts.createSourceFile(FILE, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
  return planJsxExpressionWraps(
    source, sourceFile, findStringCandidates(FILE, source, { includeJsxExpressions: true }),
  );
}

function wrapped(source: string): string {
  return applyEdits(source, plan(source).edits);
}

function reasons(source: string): string[] {
  return plan(source).residue.map((entry) => entry.reason).toSorted();
}

/** Wraps a fragment in a component, so the literal is evaluated at render like the real ones. */
function inComponent(body: string): string {
  return `export function Sample() {\n  return (\n    ${body}\n  )\n}\n`;
}

describe("planJsxExpressionWraps: what it wraps", () => {
  it("wraps both branches of a ternary rendered as a child", () => {
    assert.equal(
      wrapped(inComponent("<p>{busy ? 'Enabling...' : 'Always approve'}</p>")),
      inComponent("<p>{busy ? t`Enabling...` : t`Always approve`}</p>"),
    );
  });

  it("wraps a ternary in a whitelisted attribute", () => {
    assert.equal(
      wrapped(inComponent("<button title={busy ? 'Saving now' : 'Save now'} />")),
      inComponent("<button title={t`Saving now`} />").replace(
        "title={t`Saving now`}", "title={busy ? t`Saving now` : t`Save now`}"),
    );
  });

  it("wraps every branch of a nested chain", () => {
    assert.equal(
      wrapped(inComponent("<p>{a ? 'First one' : b ? 'Second one' : 'Third one'}</p>")),
      inComponent("<p>{a ? t`First one` : b ? t`Second one` : t`Third one`}</p>"),
    );
  });

  it("wraps a `||` default and a `&&` value", () => {
    assert.equal(
      wrapped(inComponent("<p>{title || 'Untitled workspace'}</p>")),
      inComponent("<p>{title || t`Untitled workspace`}</p>"),
    );
    assert.equal(
      wrapped(inComponent("<p>{failed && 'Something went wrong'}</p>")),
      inComponent("<p>{failed && t`Something went wrong`}</p>"),
    );
  });

  it("writes the value the literal says, not the escapes it was typed with", () => {
    assert.equal(
      wrapped(inComponent("<p>{ok ? 'It\\'s ready' : 'Not yet'}</p>")),
      inComponent("<p>{ok ? t`It's ready` : t`Not yet`}</p>"),
    );
  });

  it("tags a template rather than retyping it, so every byte survives", () => {
    assert.equal(
      wrapped(inComponent("<p>{ok ? `All done` : `Add ${vendor.name}`}</p>")),
      inComponent("<p>{ok ? t`All done` : t`Add ${vendor.name}`}</p>"),
    );
  });

  it("leaves a sibling icon out of the message", () => {
    for (const sibling of ["<Icon />", "{busy ? <Spinner /> : <Icon />}", "{/* note */}"]) {
      const source = inComponent(`<button>${sibling}{busy ? 'Saving now' : 'Save now'}</button>`);
      assert.equal(
        wrapped(source),
        inComponent(`<button>${sibling}{busy ? t\`Saving now\` : t\`Save now\`}</button>`),
        sibling,
      );
    }
  });

  it("ignores candidates of other kinds", () => {
    assert.deepEqual(plan(inComponent('<p title="Close this">Delete this workspace?</p>')).edits, []);
  });

  it("leaves nothing to do on its own output", () => {
    const once = wrapped(inComponent("<p>{busy ? 'Enabling...' : 'Always approve'}</p>"));
    assert.deepEqual(plan(once).edits, []);
    assert.equal(wrapped(once), once);
  });
});

describe("planJsxExpressionWraps: fragments", () => {
  it("refuses a literal that a template embeds in a longer sentence", () => {
    const source = inComponent("<p>{`Kept ${n} of ${ok ? 'all messages' : 'some messages'}`}</p>");
    assert.deepEqual(plan(source).edits, []);
    assert.deepEqual(reasons(source),
      ["expression-complex-template", "expression-fragment", "expression-fragment"]);
  });

  it("groups the pieces of one template together", () => {
    const source = inComponent("<p>{`Kept ${n} of ${ok ? 'all messages' : 'some messages'}`}</p>");
    const pieces = plan(source).residue.filter((entry) => entry.reason === "expression-fragment");
    assert.equal(new Set(pieces.map((entry) => entry.group)).size, 1);
  });

  it("refuses an operand of a concatenation", () => {
    const source = inComponent("<p>{'Removes it from ' + name + ' permanently now.'}</p>");
    assert.deepEqual(plan(source).edits, []);
    assert.deepEqual(reasons(source), ["expression-fragment", "expression-fragment"]);
    assert.equal(new Set(plan(source).residue.map((entry) => entry.group)).size, 1);
  });

  it("refuses a branch whose sentence continues in the sibling text", () => {
    const source = inComponent("<p>Shared by {ok ? 'one person' : 'several people'}</p>");
    assert.deepEqual(plan(source).edits, []);
    assert.deepEqual(reasons(source), ["expression-fragment", "expression-fragment"]);
  });

  it("refuses a branch beside an expression that may render words", () => {
    const source = inComponent("<p>{author.name}{merged ? 'accepted changes' : 'discarded changes'}</p>");
    assert.deepEqual(plan(source).edits, []);
    assert.deepEqual(reasons(source), ["expression-fragment", "expression-fragment"]);
  });

  it("refuses a branch beside an element that renders words", () => {
    const source = inComponent("<p><strong>Quick model:</strong>{set ? 'is chosen' : 'none set.'}</p>");
    assert.deepEqual(plan(source).edits, []);
    assert.deepEqual(reasons(source), ["expression-fragment", "expression-fragment"]);
  });
});

describe("planJsxExpressionWraps: selectors that are not just a choice of message", () => {
  it("refuses a ternary that is really a plural", () => {
    for (const condition of ["n === 1", "n !== 1", "count > 1", "items.length === 0"]) {
      const source = inComponent(`<p>{${condition} ? 'person loses' : 'people lose'}</p>`);
      assert.deepEqual(plan(source).edits, [], condition);
      assert.deepEqual(reasons(source),
        ["expression-plural-selector", "expression-plural-selector"], condition);
    }
  });

  it("sees the count in one branch of a longer chain", () => {
    const source = inComponent("<p>{empty ? 'Nothing here' : n === 1 ? 'one entry' : 'many entries'}</p>");
    assert.deepEqual(plan(source).edits, []);
    assert.equal(new Set(plan(source).residue.map((entry) => entry.reason)).size, 1);
  });

  it("wraps a ternary whose condition is not a count", () => {
    assert.equal(
      wrapped(inComponent("<p>{mode === 'edit' ? 'Edit blueprint' : 'Create blueprint'}</p>")),
      inComponent("<p>{mode === 'edit' ? t`Edit blueprint` : t`Create blueprint`}</p>"),
    );
  });

  it("refuses the whole selector when one branch holds text the detector did not report", () => {
    // "Off" is too short to be a label word, so wrapping only "Ready to go" would leave it bare and
    // invisible: the guard reports what the detector finds, and it did not find this one.
    const source = inComponent("<p>{ok ? 'Ready to go' : 'Off'}</p>");
    assert.deepEqual(plan(source).edits, []);
    assert.deepEqual(reasons(source), ["expression-partial-selector"]);
  });

  it("wraps a selector whole or not at all", () => {
    // The template branch is refused, so the branch beside it stays put too: a half-localized
    // ternary reads as finished, and the residue entry would not mention its wrapped sibling.
    const source = inComponent("<p>{ok ? 'Ready to go' : `Add ${vendorName()}`}</p>");
    assert.deepEqual(plan(source).edits, []);
    assert.deepEqual(reasons(source),
      ["expression-complex-template", "expression-partial-selector"]);
  });

  it("puts every branch of a refused selector in one group", () => {
    const source = inComponent("<p>{ok ? 'Ready to go' : `Add ${vendorName()}`}</p>");
    assert.equal(new Set(plan(source).residue.map((entry) => entry.group)).size, 1);
  });

  it("does not count an empty branch as uncovered text", () => {
    assert.equal(
      wrapped(inComponent("<p>{ok ? 'Ready to go' : ''}</p>")),
      inComponent("<p>{ok ? t`Ready to go` : ''}</p>"),
    );
  });
});

describe("planJsxExpressionWraps: values it cannot follow", () => {
  it("refuses a literal stored in a local first", () => {
    const source = inComponent(
      "<div>{items.map(() => { const label = ok ? 'Ready to go' : 'Not ready yet'\n" +
      "      return <p>{label}</p> })}</div>",
    );
    assert.deepEqual(plan(source).edits, []);
    assert.deepEqual(reasons(source),
      ["expression-indirect-value", "expression-indirect-value"]);
  });

  it("refuses a property of a record, whose siblings may be discriminants", () => {
    const source = inComponent(
      "<div>{[{ value: 'disabled', label: 'Off for everyone' }]" +
      ".map((o) => <b key={o.value}>{o.label}</b>)}</div>",
    );
    assert.deepEqual(plan(source).edits, []);
    assert.deepEqual(reasons(source), ["expression-indirect-value"]);
  });

  it("refuses an argument handed to a call", () => {
    const source = inComponent("<p>{format('Removes it from the list')}</p>");
    assert.deepEqual(plan(source).edits, []);
    assert.deepEqual(reasons(source), ["expression-indirect-value"]);
  });

  it("refuses a template whose substitutions are more than value references", () => {
    const source = inComponent("<p>{`Add ${vendorName()}`}</p>");
    assert.deepEqual(plan(source).edits, []);
    assert.deepEqual(reasons(source), ["expression-complex-template"]);
  });

  it("accepts a property-access substitution, including a non-null assertion", () => {
    assert.equal(
      wrapped(inComponent("<p>{`Shown under ${format!.output.plural} on Outputs`}</p>")),
      inComponent("<p>{t`Shown under ${format!.output.plural} on Outputs`}</p>"),
    );
  });
});

describe("planJsxExpressionWraps: values a template literal would change", () => {
  it("refuses a value that would not survive being retyped", () => {
    for (const value of [
      "Braces {like} these",
      "Use a \\`backtick\\` here",
      "Line one\\nLine two",
      "A back\\\\slash here",
    ]) {
      const source = inComponent(`<p>{title || '${value}'}</p>`);
      assert.deepEqual(plan(source).edits, [], value);
      assert.deepEqual(reasons(source), ["expression-unsafe-value"], value);
    }
  });

  it("keeps an ampersand, which a JavaScript literal does not decode", () => {
    assert.equal(
      wrapped(inComponent("<p>{title || 'Tom &amp; Jerry drink tea'}</p>")),
      inComponent("<p>{title || t`Tom &amp; Jerry drink tea`}</p>"),
    );
  });
});

describe("planJsxExpressionWraps: module scope", () => {
  it("refuses JSX built before a locale can be active", () => {
    const source = "const Banner = <p title={ok ? 'All good here' : 'Something failed'} />\n";
    assert.deepEqual(plan(source).edits, []);
    assert.deepEqual(reasons(source),
      ["expression-module-scope", "expression-module-scope"]);
  });

  it("wraps the same JSX inside a component", () => {
    assert.equal(
      wrapped(inComponent("<p title={ok ? 'All good here' : 'Something failed'} />")),
      inComponent("<p title={ok ? t`All good here` : t`Something failed`} />"),
    );
  });
});

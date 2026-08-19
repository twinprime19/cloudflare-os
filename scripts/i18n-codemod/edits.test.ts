import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { applyEdits, emptyPlan, type Edit } from "./edits.ts";

describe("applyEdits", () => {
  const source = "alpha beta gamma";

  it("returns the source unchanged when there is nothing to do", () => {
    assert.equal(applyEdits(source, []), source);
  });

  it("applies edits regardless of the order they are given in", () => {
    const edits: Edit[] = [
      { start: 11, end: 16, replacement: "GAMMA" },
      { start: 0, end: 5, replacement: "ALPHA" },
    ];
    assert.equal(applyEdits(source, edits), "ALPHA beta GAMMA");
    assert.equal(applyEdits(source, edits.toReversed()), "ALPHA beta GAMMA");
  });

  it("applies an insertion, which is a replacement of nothing", () => {
    assert.equal(applyEdits(source, [{ start: 5, end: 5, replacement: "!" }]), "alpha! beta gamma");
  });

  it("keeps offsets valid across an insertion that grows the file", () => {
    assert.equal(
      applyEdits(source, [
        { start: 0, end: 0, replacement: "// header\n" },
        { start: 11, end: 16, replacement: "GAMMA" },
      ]),
      "// header\nalpha beta GAMMA",
    );
  });

  it("refuses overlapping edits rather than producing a mangled file", () => {
    assert.throws(() => applyEdits(source, [
      { start: 0, end: 7, replacement: "x" },
      { start: 5, end: 10, replacement: "y" },
    ]), /Overlapping/);
  });

  it("refuses an inverted span", () => {
    assert.throws(() => applyEdits(source, [{ start: 7, end: 3, replacement: "x" }]), /Inverted/);
  });
});

describe("emptyPlan", () => {
  it("is empty, and is a fresh object each time", () => {
    const first = emptyPlan();
    first.residue.push({ kind: "jsx-text", text: "x", line: 1, column: 1, reason: "interpolated" });
    assert.equal(emptyPlan().residue.length, 0);
    assert.equal(emptyPlan().edits.length, 0);
  });
});

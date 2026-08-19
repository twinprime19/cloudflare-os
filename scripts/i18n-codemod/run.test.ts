import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { transformFile } from "./run.ts";

const FILE = "packages/workshop-frontend/src/Sample.tsx";

const COMPONENT = `import { Dialog } from '@cloudflare/kumo'
import { Row } from './Row'

export function Sample({ name, onClose }: { name: string; onClose: () => void }) {
  return (
    <Dialog>
      <Dialog.Title>Delete this workspace?</Dialog.Title>
      <p className="note">
        Everything in {name} goes with it.
      </p>
      <p className="hint">
        Read the <a href="/docs">documentation</a>
      </p>
      <input placeholder="Type the workspace name" />
      <button aria-label="Close" onClick={onClose}>Cancel</button>
      <Row />
    </Dialog>
  )
}
`;

describe("transformFile", () => {
  const result = transformFile(FILE, COMPONENT);

  it("wraps what it can and adds the imports those wraps need", () => {
    assert.ok(result.output);
    assert.match(result.output, /import \{ t \} from '@lingui\/core\/macro'/);
    assert.match(result.output, /import \{ Trans \} from '@lingui\/react\/macro'/);
    assert.match(result.output, /<Dialog\.Title><Trans>Delete this workspace\?<\/Trans><\/Dialog\.Title>/);
    assert.match(result.output, /placeholder=\{t`Type the workspace name`\}/);
    assert.match(result.output, /aria-label=\{t`Close`\}/);
    assert.match(result.output, /<Trans>Cancel<\/Trans>/);
  });

  it("puts the imports where the file's own package imports end", () => {
    assert.ok(result.output);
    assert.match(
      result.output,
      /'@cloudflare\/kumo'\nimport \{ t \} from '@lingui\/core\/macro'\nimport \{ Trans \} from '@lingui\/react\/macro'\nimport \{ Row \}/,
    );
  });

  it("reports the sentence broken by an interpolation instead of wrapping its halves", () => {
    assert.ok(result.output);
    assert.doesNotMatch(result.output, /<Trans>Everything in<\/Trans>/);
    const reported = result.residue.filter((entry) => entry.reason === "interpolated");
    assert.deepEqual(reported.map((entry) => entry.text), ["Everything in", "goes with it."]);
    assert.equal(new Set(reported.map((entry) => entry.group)).size, 1);
  });

  it("reports the sentence split across markup instead of wrapping its start", () => {
    assert.ok(result.output);
    assert.doesNotMatch(result.output, /<Trans>Read the<\/Trans>/);
    assert.deepEqual(
      result.residue.filter((entry) => entry.reason === "split-by-markup")
        .map((entry) => entry.text),
      ["Read the"],
    );
  });

  it("changes nothing outside the wraps and the import line", () => {
    assert.ok(result.output);
    const stripped = result.output
      .replaceAll(/<\/?Trans>/g, "")
      .replaceAll(/\{t`([^`]*)`\}/g, '"$1"')
      .replace("import { t } from '@lingui/core/macro'\n", "")
      .replace("import { Trans } from '@lingui/react/macro'\n", "");
    assert.equal(stripped, COMPONENT);
  });

  it("reports residue at the line it ends up on, not the line it started on", () => {
    assert.ok(result.output);
    // Two import lines were added above everything, so every reported line moved with them. A
    // reader opening the file at the reported line must find the string there.
    const lines = result.output.split("\n");
    assert.ok(result.residue.length > 0);
    for (const entry of result.residue) {
      assert.ok(
        lines[entry.line - 1]?.includes(entry.text),
        `"${entry.text}" is not on output line ${entry.line}: ${lines[entry.line - 1]}`,
      );
      assert.equal(
        lines[entry.line - 1]?.slice(entry.column - 1, entry.column - 1 + entry.text.length),
        entry.text,
      );
    }
  });

  it("is a no-op the second time", () => {
    assert.ok(result.output);
    const again = transformFile(FILE, result.output);
    assert.deepEqual(again.edits, []);
    assert.equal(again.output, undefined);
  });

  it("counts every detected string, wrapped or not", () => {
    assert.equal(result.candidates, result.edits.filter((edit) => edit.kind !== "import").length +
      result.residue.length);
  });
});

describe("transformFile: macro name conflicts", () => {
  const source = `export function Sample({ items }: { items: string[] }) {
  const labels = items.map((t) => t.trim())
  return <button aria-label="Close" title={labels[0]}>Cancel</button>
}
`;
  const result = transformFile(FILE, source);

  it("still wraps JSX text when only `t` is taken", () => {
    assert.ok(result.output);
    assert.match(result.output, /<Trans>Cancel<\/Trans>/);
    assert.match(result.output, /import \{ Trans \} from '@lingui\/react\/macro'/);
  });

  it("leaves the attribute alone and says why", () => {
    assert.ok(result.output);
    assert.match(result.output, /aria-label="Close"/);
    assert.deepEqual(result.residue.map((entry) => entry.reason), ["identifier-conflict"]);
  });
});

describe("transformFile: JSX expression literals", () => {
  const source = `export function Sample({ busy, name }: { busy: boolean; name: string }) {
  return (
    <div>
      <button title={busy ? 'Enabling...' : 'Always approve'}>
        {busy ? 'Enabling...' : 'Always approve'}
      </button>
      <p>Shared by {busy ? 'one person' : 'several people'}</p>
      <span>{name}</span>
    </div>
  )
}
`;

  it("leaves them alone when the detector is asked not to look", () => {
    const result = transformFile(FILE, source, { detect: { includeJsxExpressions: false } });
    assert.equal(result.output, undefined);
    assert.deepEqual(result.residue.map((entry) => entry.kind), ["jsx-text"]);
  });

  it("wraps them and adds only the `t` import", () => {
    const result = transformFile(FILE, source);
    assert.ok(result.output);
    assert.match(result.output, /title=\{busy \? t`Enabling\.\.\.` : t`Always approve`\}/);
    assert.match(result.output, /\{busy \? t`Enabling\.\.\.` : t`Always approve`\}\n/);
    assert.match(result.output, /import \{ t \} from '@lingui\/core\/macro'/);
    assert.doesNotMatch(result.output, /@lingui\/react\/macro/);
  });

  it("reports the branches whose sentence continues in the sibling text", () => {
    const result = transformFile(FILE, source);
    assert.ok(result.output);
    assert.match(result.output, /Shared by \{busy \? 'one person' : 'several people'\}/);
    assert.deepEqual(
      result.residue.map((entry) => entry.reason).toSorted(),
      // The JSX text half of the same sentence is refused too, by the rule that already existed.
      ["expression-fragment", "expression-fragment", "interpolated"],
    );
  });

  it("reports residue at the line it ends up on", () => {
    const result = transformFile(FILE, source);
    assert.ok(result.output);
    const lines = result.output.split("\n");
    for (const entry of result.residue) {
      assert.ok(lines[entry.line - 1]?.includes(entry.text), entry.text);
    }
  });

  it("is a no-op the second time", () => {
    const once = transformFile(FILE, source);
    assert.ok(once.output);
    const again = transformFile(FILE, once.output);
    assert.deepEqual(again.edits, []);
    assert.equal(again.output, undefined);
  });

  it("leaves them alone when the file already binds `t`", () => {
    const conflicting = `export function Sample({ items, busy }: { items: string[]; busy: boolean }) {
  const labels = items.map((t) => t.trim())
  return <p title={labels[0]}>{busy ? 'Enabling...' : 'Always approve'}</p>
}
`;
    const result = transformFile(FILE, conflicting);
    assert.equal(result.output, undefined);
    assert.deepEqual(new Set(result.residue.map((entry) => entry.reason)),
      new Set(["identifier-conflict"]));
  });
});

describe("transformFile: file kinds", () => {
  it("refuses a .ts file rather than counting candidates it cannot act on", () => {
    // `detect.ts` reports `string-literal` candidates there, and no transform here plans an edit
    // for one, which would silently break "candidates == edits + residue".
    assert.throws(
      () => transformFile("packages/workshop-frontend/src/sample.ts",
        "throw new Error('Choose an image file.')"),
      /transforms \.tsx files only/,
    );
  });
});

describe("transformFile: nothing to do", () => {
  it("returns no output for a file with no candidates", () => {
    const result = transformFile(FILE, "export const Empty = () => <p>{value}</p>\n");
    assert.equal(result.output, undefined);
    assert.deepEqual(result.edits, []);
    assert.deepEqual(result.residue, []);
    assert.equal(result.candidates, 0);
  });

  it("returns no output for a file that is already wrapped", () => {
    const source = "import { Trans } from '@lingui/react/macro'\n" +
      "export const Done = () => <p><Trans>All set</Trans></p>\n";
    assert.equal(transformFile(FILE, source).output, undefined);
  });
});

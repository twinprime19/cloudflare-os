import assert from "node:assert/strict";
import { describe, it } from "node:test";
import ts from "typescript6";
import { applyEdits } from "./edits.ts";
import { conflictingBindings, detectImportStyle, planMacroImports } from "./imports.ts";

const FILE = "packages/workshop-frontend/src/Sample.tsx";
const BOTH = { trans: true, t: true };

function parse(source: string): ts.SourceFile {
  return ts.createSourceFile(FILE, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
}

function added(source: string, needs = BOTH): string {
  return applyEdits(source, planMacroImports(source, parse(source), needs).edits);
}

describe("planMacroImports", () => {
  it("adds nothing when nothing is needed", () => {
    const source = "import { Button } from '@cloudflare/kumo'\n";
    assert.equal(added(source, { trans: false, t: false }), source);
  });

  it("places the macros after the package imports and before the relative ones", () => {
    assert.equal(
      added("import { Dialog } from '@cloudflare/kumo'\nimport { Row } from './Row'\n"),
      "import { Dialog } from '@cloudflare/kumo'\n" +
      "import { t } from '@lingui/core/macro'\n" +
      "import { Trans } from '@lingui/react/macro'\n" +
      "import { Row } from './Row'\n",
    );
  });

  it("adds only the macro that is needed", () => {
    assert.equal(
      added("import { Dialog } from '@cloudflare/kumo'\n", { trans: true, t: false }),
      "import { Dialog } from '@cloudflare/kumo'\nimport { Trans } from '@lingui/react/macro'\n",
    );
  });

  it("follows the file's quote and semicolon style", () => {
    assert.equal(
      added('import { Dialog } from "@cloudflare/kumo";\n', { trans: true, t: false }),
      'import { Dialog } from "@cloudflare/kumo";\nimport { Trans } from "@lingui/react/macro";\n',
    );
  });

  it("merges into an existing import of the same module", () => {
    assert.equal(
      added("import { Plural } from '@lingui/react/macro'\n", { trans: true, t: false }),
      "import { Plural, Trans } from '@lingui/react/macro'\n",
    );
  });

  it("adds its own declaration rather than touching a type-only import", () => {
    assert.equal(
      added("import type { TransProps } from '@lingui/react/macro'\n", { trans: true, t: false }),
      "import type { TransProps } from '@lingui/react/macro'\n" +
      "import { Trans } from '@lingui/react/macro'\n",
    );
  });

  it("adds nothing when the imports are already there", () => {
    const source = "import { t } from '@lingui/core/macro'\n" +
      "import { Trans } from '@lingui/react/macro'\n";
    assert.deepEqual(planMacroImports(source, parse(source), BOTH).edits, []);
  });

  it("puts the imports at the top of a file that has none", () => {
    assert.equal(
      added("export const Empty = () => <p>Nothing yet</p>\n", { trans: true, t: false }),
      "import { Trans } from '@lingui/react/macro'\n\nexport const Empty = () => <p>Nothing yet</p>\n",
    );
  });

  it("falls back to the last import when every import is relative", () => {
    assert.equal(
      added("import { Row } from './Row'\n", { trans: true, t: false }),
      "import { Row } from './Row'\nimport { Trans } from '@lingui/react/macro'\n",
    );
  });
});

describe("detectImportStyle", () => {
  it("defaults to the frontend's style when there is nothing to copy", () => {
    assert.deepEqual(detectImportStyle("const a = 1\n", parse("const a = 1\n")),
      { quote: "'", semicolon: false });
  });
});

describe("conflictingBindings", () => {
  it("finds nothing in a file that does not use the names", () => {
    assert.deepEqual(conflictingBindings(parse("const a = 1\n")), []);
  });

  it("does not report the macro imports themselves", () => {
    const source = "import { t } from '@lingui/core/macro'\n" +
      "import { Trans } from '@lingui/react/macro'\n";
    assert.deepEqual(conflictingBindings(parse(source)), []);
  });

  it("reports a local variable that takes the name", () => {
    assert.deepEqual(conflictingBindings(parse("const t = setTimeout(f, 10)\n")), ["t"]);
  });

  it("reports a callback parameter that shadows the name", () => {
    assert.deepEqual(conflictingBindings(parse("items.map((t) => t.id)\n")), ["t"]);
  });

  it("reports the same name imported from somewhere else", () => {
    assert.deepEqual(conflictingBindings(parse("import { Trans } from 'react-i18next'\n")),
      ["Trans"]);
  });

  it("reports each name separately, so one does not block the other", () => {
    const source = "import { Trans } from '@lingui/react/macro'\nconst t = 1\n";
    assert.deepEqual(conflictingBindings(parse(source)), ["t"]);
  });
});

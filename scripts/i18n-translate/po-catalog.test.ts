import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, it } from "node:test";
import {
  escapePoString,
  parsePo,
  serializePo,
  setMsgstr,
  unescapePoString,
} from "./po-catalog.ts";

const REPO_ROOT = join(import.meta.dirname, "..", "..");
const EN_PO = join(REPO_ROOT, "packages", "i18n", "locales", "en.po");

describe("unescapePoString / escapePoString", () => {
  it("round-trips every escape lingui writes", () => {
    const logical = 'a "quoted" line\nwith\ttab and \\backslash';
    assert.equal(unescapePoString(escapePoString(logical)), logical);
  });

  it("unescapes the shapes found in catalogs", () => {
    assert.equal(unescapePoString('ship\\n'), "ship\n");
    assert.equal(unescapePoString('say \\"hi\\"'), 'say "hi"');
  });

  it("throws on an escape lingui never writes", () => {
    assert.throws(() => unescapePoString("\\x41"), /Unknown PO escape/);
    assert.throws(() => unescapePoString("dangling\\"), /Unknown PO escape/);
  });
});

describe("parsePo against the real catalog", () => {
  const source = readFileSync(EN_PO, "utf8");
  const catalog = parsePo(source);

  it("serializes byte-identically", () => {
    assert.equal(serializePo(catalog), source);
  });

  it("finds all 877 messages", () => {
    assert.equal(catalog.entries.length, 877);
  });

  it("joins continuation lines into one logical string", () => {
    const multiline = catalog.entries.find((entry) => entry.msgid.includes("ACME Corp"));
    assert.ok(multiline, "the known multi-line entry exists");
    assert.ok(multiline.msgid.includes("ship\ninternationally"));
    assert.equal(multiline.msgidLines.length, 2);
  });

  it("keeps placeholder comments verbatim", () => {
    const commented = catalog.entries.find((entry) => entry.comments.length > 0);
    assert.ok(commented);
    assert.match(commented.comments[0], /^#\. placeholder /);
  });
});

describe("parsePo failure shapes", () => {
  it("rejects a file without the header block", () => {
    assert.throws(() => parsePo('msgid "hi"\nmsgstr ""\n'), /header block/);
  });

  it("rejects unknown line shapes", () => {
    assert.throws(
      () => parsePo('msgid ""\nmsgstr ""\n\n#: origin.tsx:1\nmsgid "a"\nmsgstr ""\n'),
      /unrecognized PO line shape/,
    );
  });

  it("rejects a file without a trailing newline", () => {
    assert.throws(() => parsePo('msgid ""\nmsgstr ""'), /end with a newline/);
  });

  it("rejects duplicate msgids", () => {
    assert.throws(
      () => parsePo('msgid ""\nmsgstr ""\n\nmsgid "a"\nmsgstr ""\n\nmsgid "a"\nmsgstr ""\n'),
      /Duplicate msgid/,
    );
  });
});

describe("setMsgstr", () => {
  it("writes a single-line translation on the keyword line", () => {
    const catalog = parsePo('msgid ""\nmsgstr ""\n\nmsgid "Cancel"\nmsgstr ""\n');
    setMsgstr(catalog.entries[0], "Hủy");
    assert.equal(
      serializePo(catalog),
      'msgid ""\nmsgstr ""\n\nmsgid "Cancel"\nmsgstr "Hủy"\n',
    );
  });

  it("splits a translation with embedded newlines into continuation lines", () => {
    const catalog = parsePo('msgid ""\nmsgstr ""\n\nmsgid "a"\nmsgstr ""\n');
    setMsgstr(catalog.entries[0], "dòng một\ndòng hai");
    const roundTripped = parsePo(serializePo(catalog));
    assert.equal(roundTripped.entries[0].msgstr, "dòng một\ndòng hai");
    assert.equal(catalog.entries[0].msgstrLines.length, 2);
    assert.equal(catalog.entries[0].msgstrLines[0], 'msgstr "dòng một\\n"');
  });

  it("escapes quotes in translations", () => {
    const catalog = parsePo('msgid ""\nmsgstr ""\n\nmsgid "a"\nmsgstr ""\n');
    setMsgstr(catalog.entries[0], 'nói "chào"');
    assert.equal(parsePo(serializePo(catalog)).entries[0].msgstr, 'nói "chào"');
  });
});

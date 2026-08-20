// Minimal parser/serializer for the PO dialect `lingui extract` writes with this repo's config
// (`origins: false`, `@lingui/format-po`). It is deliberately NOT a general PO parser: it accepts
// exactly the shapes the real catalogs contain -- `#.` extracted comments, `msgid`/`msgstr`, and
// string-continuation lines -- and throws on anything else. Failing loudly is the contract: if a
// future lingui version emits a new shape, the round-trip test against the real catalog breaks
// before any tool silently mangles translations.
//
// Raw source lines are preserved per entry so that parse -> serialize is byte-identical for
// untouched entries. Only `setMsgstr` re-generates lines, and only for the entry it changes;
// `lingui extract` renormalizes formatting afterwards anyway, so newly written msgstr lines just
// need to be valid PO, not byte-matched to lingui's own chunking.
import { readFileSync, writeFileSync } from "node:fs";

/** One translatable catalog entry. The header block is held separately on {@link PoCatalog}. */
export interface PoEntry {
  /** `#.` comment lines, verbatim including the `#.` prefix. */
  comments: string[];
  /** Logical (unescaped, concatenated) msgid value. */
  msgid: string;
  /** Logical msgstr value; empty string when untranslated. */
  msgstr: string;
  /** Raw `msgid ...` line plus its continuation lines, exactly as in the source. */
  msgidLines: string[];
  /** Raw `msgstr ...` line plus its continuation lines, exactly as in the source. */
  msgstrLines: string[];
}

/** A parsed catalog: the verbatim header block plus the entries in file order. */
export interface PoCatalog {
  /** The `msgid ""` header block's raw lines, re-emitted untouched. */
  headerLines: string[];
  entries: PoEntry[];
}

/** Unescapes the inner content of one PO string literal. Throws on escapes lingui never writes. */
export function unescapePoString(raw: string): string {
  let out = "";
  for (let i = 0; i < raw.length; i++) {
    const char = raw[i];
    if (char !== "\\") {
      out += char;
      continue;
    }
    const next = raw[++i];
    if (next === "n") out += "\n";
    else if (next === "t") out += "\t";
    else if (next === "r") out += "\r";
    else if (next === '"') out += '"';
    else if (next === "\\") out += "\\";
    else throw new Error(`Unknown PO escape "\\${next ?? "<end>"}" in: ${raw}`);
  }
  return out;
}

/** Escapes a logical string into PO string-literal content. Inverse of {@link unescapePoString}. */
export function escapePoString(value: string): string {
  return value
    .replaceAll("\\", "\\\\")
    .replaceAll('"', '\\"')
    .replaceAll("\n", "\\n")
    .replaceAll("\t", "\\t")
    .replaceAll("\r", "\\r");
}

/** Extracts the quoted content of a `msgid "..."`, `msgstr "..."`, or bare `"..."` line. */
function quotedContent(line: string, lineNumber: number): string {
  const openQuote = line.indexOf('"');
  if (openQuote === -1 || !line.endsWith('"') || line.length < openQuote + 2) {
    throw new Error(`Line ${lineNumber} is not a PO string line: ${line}`);
  }
  return line.slice(openQuote + 1, -1);
}

/**
 * Splits a logical string into the chunks lingui's formatter puts on separate lines: one chunk per
 * embedded newline, the newline staying with the chunk it ends. A trailing newline does not
 * produce an empty final chunk.
 */
function chunksAfterNewlines(value: string): string[] {
  const chunks: string[] = [];
  let start = 0;
  for (let i = 0; i < value.length; i++) {
    if (value[i] === "\n") {
      chunks.push(value.slice(start, i + 1));
      start = i + 1;
    }
  }
  if (start < value.length || chunks.length === 0) chunks.push(value.slice(start));
  return chunks;
}

/** Builds the raw line group for a keyword and value, matching lingui's body-entry chunking. */
function buildStringLines(keyword: "msgid" | "msgstr", value: string): string[] {
  const chunks = chunksAfterNewlines(value);
  const lines = [`${keyword} "${escapePoString(chunks[0])}"`];
  for (const chunk of chunks.slice(1)) lines.push(`"${escapePoString(chunk)}"`);
  return lines;
}

/** Replaces an entry's translation, regenerating its raw msgstr lines. */
export function setMsgstr(entry: PoEntry, value: string): void {
  entry.msgstr = value;
  entry.msgstrLines = buildStringLines("msgstr", value);
}

/** One blank-line-separated block of raw lines. */
type Block = { lines: string[]; startLine: number };

function splitBlocks(source: string): Block[] {
  if (!source.endsWith("\n")) throw new Error("PO file does not end with a newline.");
  const lines = source.slice(0, -1).split("\n");
  const blocks: Block[] = [];
  let current: Block | undefined;
  for (let i = 0; i < lines.length; i++) {
    if (lines[i] === "") {
      current = undefined;
      continue;
    }
    if (!current) {
      current = { lines: [], startLine: i + 1 };
      blocks.push(current);
    }
    current.lines.push(lines[i]);
  }
  return blocks;
}

/** Parses one non-header block into an entry. */
function parseEntry(block: Block): PoEntry {
  const comments: string[] = [];
  const msgidLines: string[] = [];
  const msgstrLines: string[] = [];
  // Which group continuation `"..."` lines currently extend.
  let target: string[] | undefined;
  for (let i = 0; i < block.lines.length; i++) {
    const line = block.lines[i];
    const lineNumber = block.startLine + i;
    if (line.startsWith("#. ")) {
      if (target) throw new Error(`Line ${lineNumber}: comment after strings began: ${line}`);
      comments.push(line);
    } else if (line.startsWith("msgid ")) {
      if (msgidLines.length > 0) throw new Error(`Line ${lineNumber}: second msgid in one entry.`);
      msgidLines.push(line);
      target = msgidLines;
    } else if (line.startsWith("msgstr ")) {
      if (msgstrLines.length > 0 || msgidLines.length === 0) {
        throw new Error(`Line ${lineNumber}: msgstr out of order.`);
      }
      msgstrLines.push(line);
      target = msgstrLines;
    } else if (line.startsWith('"')) {
      if (!target) throw new Error(`Line ${lineNumber}: continuation before msgid: ${line}`);
      target.push(line);
    } else {
      throw new Error(`Line ${lineNumber}: unrecognized PO line shape: ${line}`);
    }
  }
  if (msgidLines.length === 0 || msgstrLines.length === 0) {
    throw new Error(`Entry at line ${block.startLine} is missing msgid or msgstr.`);
  }
  const logical = (group: string[], startAt: number): string =>
    group.map((line, i) => unescapePoString(quotedContent(line, startAt + i))).join("");
  return {
    comments,
    msgid: logical(msgidLines, block.startLine + comments.length),
    msgstr: logical(msgstrLines, block.startLine + comments.length + msgidLines.length),
    msgidLines,
    msgstrLines,
  };
}

/** Parses catalog source text. The first block must be the `msgid ""` header. */
export function parsePo(source: string): PoCatalog {
  const blocks = splitBlocks(source);
  if (blocks.length === 0 || blocks[0].lines[0] !== 'msgid ""') {
    throw new Error('PO file does not start with a `msgid ""` header block.');
  }
  const entries = blocks.slice(1).map(parseEntry);
  const seen = new Set<string>();
  for (const entry of entries) {
    if (seen.has(entry.msgid)) throw new Error(`Duplicate msgid: ${entry.msgid}`);
    seen.add(entry.msgid);
  }
  return { headerLines: blocks[0].lines, entries };
}

/** Serializes back to catalog text. Byte-identical to the input for untouched catalogs. */
export function serializePo(catalog: PoCatalog): string {
  const blocks = [catalog.headerLines.join("\n")];
  for (const entry of catalog.entries) {
    blocks.push([...entry.comments, ...entry.msgidLines, ...entry.msgstrLines].join("\n"));
  }
  return blocks.join("\n\n") + "\n";
}

/** Reads and parses a catalog file. */
export function readPoFile(path: string): PoCatalog {
  return parsePo(readFileSync(path, "utf8"));
}

/** Serializes and writes a catalog file. */
export function writePoFile(path: string, catalog: PoCatalog): void {
  writeFileSync(path, serializePo(catalog));
}

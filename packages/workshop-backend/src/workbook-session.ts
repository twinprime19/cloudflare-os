// The programmatic Workbook binding: the RpcTarget behind a spreadsheet attachment's env entry in
// the agent's executeCode sandbox (workbook-binding.d.ts defines the agent-facing contract).
//
// A session is minted by the overseer's binding-loopback dispatch (startGatekeeperSession, target
// type "workbook"). What minting checks is the caller and the chat: the loopback is built for the
// agent's own executeCode, so a gadget or a user calling it directly is refused, as is an
// attachment belonging to another chat. Past that the session is an ordinary binding in the
// sandbox's env -- the agent's code can hand the stub to whatever it calls, and what that holder
// can read is this one committed attachment of this one chat, read-only.
//
// Unlike the worktree binding, none of what it serves is turn state: an attachment's rows are
// written once, when the upload is staged, and are never edited. So the session is a pure reader
// over durable records and is deliberately not scoped to the executeCode run that minted it --
// there is no live turn for a retained stub to revive against, only the same immutable rows.
//
// Rows live in one record per chunk, indexed by WorkbookMeta.chunkRowStarts: chunk i holds the
// sheet's rows from chunkRowStarts[i] up to chunkRowStarts[i + 1] (or the sheet's rowCount, for
// the last chunk). Reading a range therefore decodes only the pages that overlap it, which is what
// keeps a page of a 200,000-row sheet as cheap as a page of a small one.

import { RpcTarget } from "cloudflare:workers";
import { validateRpc } from "capnweb-validate";
import type { CellValue, Workbook } from "./workbook-binding";
import { formatNumber, type WorkbookMeta } from "./chat-attachment-workbook";
import { columnLabel, parseColumnLabel, renderAddressedRows } from "./workbook-grid";

/**
 * Ceiling on the rows one getRows() call returns.
 *
 * Wide of any realistic single read while bounding both the RPC payload (20,000 rows of 10 columns
 * is a few MB of JSON) and the memory the decode holds. A request past it fails with instructions
 * to page rather than returning a silently truncated answer the agent would go on to compute over.
 */
export const MAX_WORKBOOK_ROWS_PER_CALL = 20000;

/**
 * What the session needs from the overseer: one stored row page, by attachment, sheet and chunk.
 * Structurally satisfied by OverseerImpl, which owns the record layout and its keys.
 */
export interface WorkbookSessionHost {
  readWorkbookChunk(fileId: string, sheetIndex: number, chunkIndex: number): Uint8Array | undefined;
}

/**
 * The Workbook binding served to executeCode. One instance per (execution, attachment); see the
 * module doc for why it holds no per-turn state.
 */
@validateRpc()
export class WorkbookSessionImpl extends RpcTarget implements Workbook {
  constructor(private host: WorkbookSessionHost, private fileId: string,
              private meta: WorkbookMeta) {
    super();
  }

  async listSheets(): Promise<{name: string, rowCount: number, colCount: number}[]> {
    return this.meta.sheets.map(({name, rowCount, colCount}) => ({name, rowCount, colCount}));
  }

  async getRows(sheet: string, start?: number, end?: number): Promise<CellValue[][]> {
    let sheetIndex = findSheetIndex(this.meta, sheet);
    let {rowCount} = this.meta.sheets[sheetIndex];

    let from = start ?? 0;
    let requestedEnd = end ?? rowCount;
    if (!Number.isInteger(from) || !Number.isInteger(requestedEnd)) {
      throw new Error("start and end must be whole numbers of rows.");
    }
    if (from < 0 || requestedEnd < 0) {
      throw new Error("start and end must be 0 or greater; rows are counted from 0.");
    }
    // An end past the sheet is clamped rather than refused: asking for a full page off the end of
    // a sheet is how paging through one naturally finishes.
    let to = Math.min(requestedEnd, rowCount);
    // Likewise a start past the end returns nothing at all, which is the signal a paging loop
    // stops on.
    if (from >= rowCount) return [];
    if (to < from) {
      throw new Error(`end (${requestedEnd}) is before start (${from}); the range is [start, end).`);
    }
    if (to - from > MAX_WORKBOOK_ROWS_PER_CALL) {
      throw new Error(
          `That range is ${to - from} rows, and at most ${MAX_WORKBOOK_ROWS_PER_CALL} can be ` +
          `read per call. Read the sheet in pages, e.g. getRows(${JSON.stringify(sheet)}, ` +
          `${from}, ${from + MAX_WORKBOOK_ROWS_PER_CALL}) and then from there.`);
    }

    return readSheetRows(this.host, this.fileId, sheetIndex, this.meta, from, to);
  }
}

/**
 * The index of the sheet named `sheet`, or an error naming the sheets the workbook does have.
 * Sheet names are unique within a workbook (spreadsheet software enforces it), so the first match
 * is the only match.
 */
function findSheetIndex(meta: WorkbookMeta, sheet: string): number {
  let sheetIndex = meta.sheets.findIndex(entry => entry.name === sheet);
  if (sheetIndex < 0) {
    let names = meta.sheets.map(entry => JSON.stringify(entry.name)).join(", ");
    throw new Error(
        `This workbook has no sheet named ${JSON.stringify(sheet)}. Its sheets are: ${names}.`);
  }
  return sheetIndex;
}

/**
 * Rows `[from, to)` of one sheet (0-based), decoding only the stored pages that overlap them. The
 * one reader behind both getRows and readSheet, so the two can never disagree about what a row
 * holds. The caller validates the range; this applies no cap of its own.
 */
export function readSheetRows(host: WorkbookSessionHost, fileId: string, sheetIndex: number,
                              meta: WorkbookMeta, from: number, to: number): CellValue[][] {
  let {name, rowCount, chunkRowStarts} = meta.sheets[sheetIndex];
  let rows: CellValue[][] = [];
  for (let [chunkIndex, chunkStart] of chunkRowStarts.entries()) {
    let chunkEnd = chunkRowStarts[chunkIndex + 1] ?? rowCount;
    if (chunkEnd <= from) continue;
    if (chunkStart >= to) break;
    let stored = host.readWorkbookChunk(fileId, sheetIndex, chunkIndex);
    if (stored === undefined) {
      // The index and the pages are written in one transaction and deleted in one transaction,
      // so a gap means the records were lost, not that this range is empty. Say so rather than
      // handing back a short answer the agent would treat as the data.
      throw new Error(
          `The stored data for sheet ${JSON.stringify(name)} is incomplete; this workbook ` +
          `cannot be read.`);
    }
    let chunkRows = JSON.parse(new TextDecoder().decode(stored)) as CellValue[][];
    let sliceStart = Math.max(from - chunkStart, 0);
    let sliceEnd = Math.min(to - chunkStart, chunkRows.length);
    for (let index = sliceStart; index < sliceEnd; index++) rows.push(chunkRows[index]);
  }
  return rows;
}

// ---------------------------------------------------------------------------------------------
// readSheet: the agent's labeled read path, a tool rather than code. Where getRows hands arrays to
// code that computes, readSheet hands text to the model that reads -- every cell written with its
// Excel address, so a value is never located by counting positions. It speaks the spreadsheet's
// own coordinates (rows from 1, columns by letter, ranges inclusive) because that is how the
// attachment text and the user name cells.

/** Ceiling on the rows one readSheet call returns: a page the model reads, not a dataset. */
export const MAX_READ_SHEET_ROWS = 200;

/**
 * Ceiling on the rendered rows of one readSheet call. A result stays in the context of every later
 * turn, so a page is kept to a few thousand tokens; past it the text stops at a whole row and says
 * where to continue.
 */
export const MAX_READ_SHEET_BYTES = 24 * 1024;

/** The rows read when no range is given: enough to see a sheet's headers and first records. */
const DEFAULT_READ_SHEET_ROWS = 50;

/** A parsed readSheet range: 0-based, half-open, clamped to the sheet's used extent. */
export type SheetRange = {
  from: number;
  to: number;
  colStart: number;
  colEnd: number;
};

const RANGE_FORMS = `Use "A150:AV160" (cells) or "150:160" (whole rows); rows count from 1 and ` +
    `both ends are included.`;

/**
 * Parse a readSheet `range` against the sheet it reads. Accepts `A150:AV160` and `150:160`
 * (1-based, inclusive, case-insensitive); an omitted range is the first 50 rows, all columns. An
 * end past the sheet is clamped, as getRows clamps; a start past it is an error, because an empty
 * page would read as "nothing there" rather than "you asked beyond the data". More than
 * MAX_READ_SHEET_ROWS rows (after clamping) is refused with the paging call to make instead.
 */
export function parseSheetRange(range: string | undefined,
                                sheet: {rowCount: number, colCount: number}): SheetRange {
  let firstRow = 1;
  let lastRow = DEFAULT_READ_SHEET_ROWS;
  let firstCol = 0;
  let lastCol: number | undefined;

  if (range !== undefined && range.trim() !== "") {
    let text = range.trim().toUpperCase();
    let cells = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(text);
    let rows = /^(\d+):(\d+)$/.exec(text);
    if (cells) {
      try {
        firstCol = parseColumnLabel(cells[1]);
        lastCol = parseColumnLabel(cells[3]);
      } catch {
        throw new Error(`Invalid range ${JSON.stringify(range)}. ${RANGE_FORMS}`);
      }
      firstRow = Number(cells[2]);
      lastRow = Number(cells[4]);
    } else if (rows) {
      firstRow = Number(rows[1]);
      lastRow = Number(rows[2]);
    } else {
      throw new Error(`Invalid range ${JSON.stringify(range)}. ${RANGE_FORMS}`);
    }
    if (firstRow < 1 || lastRow < 1) {
      throw new Error(`Invalid range ${JSON.stringify(range)}: rows count from 1.`);
    }
    if (lastRow < firstRow || (lastCol !== undefined && lastCol < firstCol)) {
      throw new Error(
          `Invalid range ${JSON.stringify(range)}: it runs backwards. Write the top-left cell ` +
          `first, e.g. "A150:AV160".`);
    }
  }

  if (firstRow > sheet.rowCount) {
    throw new Error(`Row ${firstRow} is past the end of this sheet, which has ` +
        `${formatNumber(sheet.rowCount)} rows.`);
  }
  if (firstCol >= sheet.colCount) {
    throw new Error(`Column ${columnLabel(firstCol)} is past the last column of this sheet` +
        (sheet.colCount > 0 ? `, ${columnLabel(sheet.colCount - 1)}.` : "."));
  }

  let from = firstRow - 1;
  let to = Math.min(lastRow, sheet.rowCount);
  if (to - from > MAX_READ_SHEET_ROWS) {
    throw new Error(
        `That range is ${to - from} rows; readSheet returns at most ${MAX_READ_SHEET_ROWS} per ` +
        `call. Read it in pages, e.g. "${firstRow}:${firstRow + MAX_READ_SHEET_ROWS - 1}".`);
  }
  return {
    from,
    to,
    colStart: firstCol,
    colEnd: Math.min((lastCol ?? sheet.colCount - 1) + 1, sheet.colCount),
  };
}

/**
 * The readSheet result for one range of one sheet: a header line naming what was read, then one
 * addressed line per non-blank row with every cell spelled out (never collapsed -- this is the
 * path for reading the values a collapsed summary elided), then, when the page crossed
 * MAX_READ_SHEET_BYTES, a notice naming the range to continue with.
 *
 * Deterministic for a given attachment and arguments; the recorded text is what replays.
 */
export function readSheetRange(host: WorkbookSessionHost, fileId: string, meta: WorkbookMeta,
                               sheet: string, range: string | undefined): string {
  let sheetIndex = findSheetIndex(meta, sheet);
  let sheetMeta = meta.sheets[sheetIndex];
  let {from, to, colStart, colEnd} = parseSheetRange(range, sheetMeta);

  // Cells left of the range become empty rather than being dropped, so every label the renderer
  // writes is the cell's real column; empty cells are never written.
  let rows = readSheetRows(host, fileId, sheetIndex, meta, from, to).map(row => {
    let cells = row.slice(0, colEnd);
    for (let index = 0; index < Math.min(colStart, cells.length); index++) cells[index] = null;
    return cells;
  });
  let rendered = renderAddressedRows(rows, from, {collapse: false, maxBytes: MAX_READ_SHEET_BYTES});

  let columns = `${columnLabel(colStart)}–${columnLabel(colEnd - 1)}`;
  let lines = [
    `Sheet ${JSON.stringify(sheet)} rows ${from + 1}–${to} of ${formatNumber(sheetMeta.rowCount)}, ` +
        `columns ${columns} (blank rows omitted)`,
  ];
  if (rendered.text !== "") lines.push(rendered.text);
  if (rendered.truncated) {
    // Continue with the same columns the call asked for; a rows-only call continues rows-only.
    let next = rendered.renderedThrough + 2;
    let wholeWidth = colStart === 0 && colEnd === sheetMeta.colCount;
    let continuation = wholeWidth ? `${next}:${to}`
        : `${columnLabel(colStart)}${next}:${columnLabel(colEnd - 1)}${to}`;
    if (rendered.text === "") {
      // Not even the first non-blank row fit (blank rows before it cost nothing and are skipped):
      // continuing from it would return the same nothing forever.
      lines.push(`… row ${next} alone is over 24 KiB. Read fewer columns of it, e.g. ` +
          `"${columnLabel(colStart)}${next}:${columnLabel(Math.min(colStart + 9, colEnd - 1))}${next}".`);
    } else {
      lines.push(`… truncated after row ${next - 1} (24 KiB). Continue with range ${continuation}.`);
    }
  }
  return lines.join("\n");
}


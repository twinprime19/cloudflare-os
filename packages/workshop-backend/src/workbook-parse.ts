// The spreadsheet parser: SheetJS reads the upload, and this walks its sheets into the rows,
// summary and chunks chat-attachment-workbook.ts describes.
//
// The only module that imports SheetJS, and deliberately reachable from nowhere the workspace
// Durable Object loads. It runs in a dynamic worker of its own, loaded fresh for each upload (see
// workbook-parser-isolate.ts and worker-runtimes/workbook-parser-runtime.ts), so a workbook that
// materializes more than an isolate can hold takes that isolate down instead of the workspace.
//
// Parsing is pure and synchronous: no storage, no Durable Object, no network. The read options
// below switch off every part of parsing whose output nothing here reads.

import { Unzip, UnzipInflate } from "fflate";
import * as XLSX from "xlsx";
import {
  MAX_WORKBOOK_ROWS,
  MAX_WORKBOOK_ROW_BYTES,
  MAX_WORKBOOK_UNCOMPRESSED_BYTES,
  WORKBOOK_EMPTY_MESSAGE,
  WORKBOOK_TOO_LARGE_MESSAGE,
  WORKBOOK_UNREADABLE_MESSAGE,
  chunkSheetRows,
  renderWorkbookSummary,
  workbookRowChunks,
} from "./chat-attachment-workbook";
import type {
  CellValue, ParsedWorkbook, SheetContent, StreamedWorkbook, WorkbookMeta,
} from "./chat-attachment-workbook";
import { SPREADSHEET_MIME_TYPES } from "./workbook-names";

// Everything switched off below is parse work whose output nothing on this path reads: formatted
// text (`.w`), number formats (`.z`), formulas (`.f`, whose cached value lands in `.v` either
// way), styles and rich-text HTML. Leaving them on roughly doubles both parse time and heap.
const READ_OPTIONS: XLSX.ParsingOptions = {
  type: "array",
  dense: true,
  cellDates: true,
  cellText: false,
  cellNF: false,
  cellFormula: false,
  cellStyles: false,
  cellHTML: false,
  sheetStubs: false,
  // One row past the ceiling, so a sheet claiming a huge extent is read as the few rows it really
  // holds plus the marker SheetJS leaves when it truncates -- which is what the walk below refuses
  // on, before a row of that sheet is materialized.
  sheetRows: MAX_WORKBOOK_ROWS + 1,
};

// The local-file header every ZIP container starts with. OOXML and OpenDocument packages are both
// ZIPs; legacy .xls is an OLE2 compound file and has no ZIP header.
const ZIP_LOCAL_FILE_HEADER = [0x50, 0x4b, 0x03, 0x04];

// How much of an upload is handed to the inflater at a time. fflate expands everything it is given
// before it hands any of it back, so this is what bounds the memory the measurement below can
// spend: deflate's best ratio is about 1032:1, so a slice this size cannot expand past ~16 MiB
// however the archive was built.
const INFLATE_SLICE_BYTES = 16 * 1024;

// BIFF error codes, which is what `v` holds for an error cell. `cellText: false` means SheetJS
// never materializes the `.w` text for them, so the code is translated here rather than reaching
// the model as a bare number.
const EXCEL_ERROR_TEXT = new Map<number, string>([
  [0x00, "#NULL!"],
  [0x07, "#DIV/0!"],
  [0x0f, "#VALUE!"],
  [0x17, "#REF!"],
  [0x1d, "#NAME?"],
  [0x24, "#NUM!"],
  [0x2a, "#N/A"],
  [0x2b, "#GETTING_DATA"],
]);

/** Thrown from the inflate callback so the walk stops at the chunk that crosses the cap. */
class WorkbookTooLargeError extends Error {
  constructor() {
    super(WORKBOOK_TOO_LARGE_MESSAGE);
    this.name = "WorkbookTooLargeError";
  }
}

function isZipContainer(bytes: Uint8Array): boolean {
  return ZIP_LOCAL_FILE_HEADER.every((byte, index) => bytes[index] === byte);
}

/**
 * Refuse a workbook whose content would not fit in the isolate, by inflating it and counting.
 *
 * The archive's own headers cannot answer this. A ZIP declares each entry's uncompressed size in
 * two places -- the central directory and the entry's local header -- and nothing makes an archive
 * honour either: the parser inflates whatever the compressed stream actually holds and finds out
 * afterwards. A declared size is therefore a claim, and the only honest measure is the inflated
 * byte count.
 *
 * So the upload is inflated here and thrown away a chunk at a time, and the running total is what
 * the cap is checked against. It is fed in slices because fflate expands everything it is handed
 * before it hands any of it back: one push of the whole upload would materialize the very thing
 * this exists to refuse. SheetJS then inflates the archive again when it parses, which is the
 * price of measuring the real quantity instead of a claimed one.
 *
 * Legacy .xls is an uncompressed OLE2 container: its uncompressed size is its raw size, which the
 * caller's upload cap already bounds, so it skips this entirely.
 */
function assertWorkbookFitsInMemory(bytes: Uint8Array): void {
  if (!isZipContainer(bytes)) return;

  let inflatedBytes = 0;
  let overCap = false;
  let failure: unknown;

  let unzip = new Unzip();
  unzip.register(UnzipInflate);
  unzip.onfile = (file) => {
    file.ondata = (error, chunk) => {
      // fflate delivers a throw from here back as an error callback on the same entry, so once
      // either flag is set the walk is over and later callbacks carry nothing new.
      if (overCap || failure !== undefined) return;
      if (error) {
        failure = error;
        return;
      }
      // The chunk is counted and dropped: its length is the whole of what this needs from it.
      inflatedBytes += chunk.length;
      if (inflatedBytes > MAX_WORKBOOK_UNCOMPRESSED_BYTES) {
        overCap = true;
        // Stops the rest of the current slice, which may span several entries.
        throw new WorkbookTooLargeError();
      }
    };
    file.start();
  };

  try {
    for (let offset = 0; offset < bytes.length; offset += INFLATE_SLICE_BYTES) {
      let end = Math.min(offset + INFLATE_SLICE_BYTES, bytes.length);
      unzip.push(bytes.subarray(offset, end), end === bytes.length);
      if (overCap || failure !== undefined) break;
    }
  } catch (error) {
    if (!overCap && failure === undefined) failure = error;
  }

  if (overCap) throw new WorkbookTooLargeError();
  // A container fflate cannot walk is not a workbook, whatever its MIME type claimed -- and one
  // this cannot walk is one whose size it cannot bound. The cause is kept for logs; the message is
  // the one the user sees.
  if (failure !== undefined) throw new Error(WORKBOOK_UNREADABLE_MESSAGE, { cause: failure });
}

/**
 * Render a date cell. SheetJS materializes the cell's wall clock in the Date's UTC components, so
 * the ISO string is the spreadsheet's own value regardless of where the isolate runs. A cell whose
 * time is midnight is a plain date and loses the time half, which is what a spreadsheet date is.
 */
function formatDateCell(value: Date): CellValue {
  if (Number.isNaN(value.getTime())) return null;
  let iso = value.toISOString();
  return iso.endsWith("T00:00:00.000Z") ? iso.slice(0, 10) : iso;
}

function toCellValue(cell: XLSX.CellObject | undefined): CellValue {
  if (!cell || cell.t === "z") return null;
  let value = cell.v;
  if (value === undefined || value === null) return null;

  switch (cell.t) {
    case "b":
      return Boolean(value);
    case "n":
      return Number(value);
    case "d":
      return value instanceof Date ? formatDateCell(value) : String(value);
    case "e":
      if (cell.w) return cell.w;
      return typeof value === "number" ? EXCEL_ERROR_TEXT.get(value) ?? String(value) : String(value);
    default:
      return String(value);
  }
}

function trimTrailingNulls(row: CellValue[]): void {
  while (row.length > 0 && row[row.length - 1] === null) row.pop();
}

/**
 * Read one sheet as rows of cell values.
 *
 * Trailing empty cells and trailing empty rows are dropped: a spreadsheet's stored range routinely
 * extends past its content -- a formatted-but-cleared column, a row someone once typed in -- and
 * keeping that padding would inflate every dimension the model is told about. Interior blanks stay,
 * because table detection reads them as separators.
 */
function readSheetRows(sheet: XLSX.WorkSheet): CellValue[][] {
  // `dense: true` makes SheetJS store rows under `!data`; the fallback covers sheets that carry no
  // grid at all, such as chartsheets.
  let data = (sheet as XLSX.DenseWorkSheet)["!data"] ?? [];
  let rows: CellValue[][] = [];

  for (let rawRow of data) {
    let row: CellValue[] = [];
    // Dense rows are arrays with holes, and a row the file never wrote is a hole itself.
    if (rawRow) for (let cell of rawRow) row.push(toCellValue(cell));
    trimTrailingNulls(row);
    rows.push(row);
  }

  while (rows.length > 0 && rows[rows.length - 1].length === 0) rows.pop();
  return rows;
}

/**
 * Whether a sheet's stored extent ran past what the read took. SheetJS records the untruncated
 * range under `!fullref` when `sheetRows` cuts a read short, and leaves it off every sheet it read
 * whole -- so its presence is exactly "this sheet claims more rows than the ceiling allows".
 */
function exceedsRowCeiling(sheet: XLSX.WorkSheet): boolean {
  let fullRef: unknown = (sheet as Record<string, unknown>)["!fullref"];
  return typeof fullRef === "string" && fullRef !== sheet["!ref"];
}

function countCells(rows: CellValue[][]): number {
  let total = 0;
  for (let row of rows) {
    for (let cell of row) if (cell !== null) total++;
  }
  return total;
}

/**
 * Parse a spreadsheet attachment into its prompt summary, its index, and the row chunks to store.
 *
 * Throws with user-facing wording on every failure: the caller surfaces the message as the reason
 * the upload was refused, so nothing here leaks a parser's own error text.
 */
export function parseWorkbookAttachment(
  bytes: Uint8Array,
  mimeType: string,
  name: string,
): ParsedWorkbook {
  // The routing decision belongs to the caller; this guards against a non-spreadsheet reaching the
  // parser, which SheetJS would otherwise read as plain text and turn into a nonsense workbook.
  if (!SPREADSHEET_MIME_TYPES.has(mimeType)) throw new Error(WORKBOOK_UNREADABLE_MESSAGE);

  assertWorkbookFitsInMemory(bytes);

  let workbook: XLSX.WorkBook;
  try {
    workbook = XLSX.read(bytes, READ_OPTIONS);
  } catch (error) {
    throw new Error(WORKBOOK_UNREADABLE_MESSAGE, { cause: error });
  }

  let contents: SheetContent[] = [];
  let sheets: ParsedWorkbook["sheets"] = [];
  let metaSheets: WorkbookMeta["sheets"] = [];
  let cellCount = 0;
  let rowCount = 0;
  // One budget for the whole workbook: the rows of every sheet land in the same storage and the
  // same isolate, so they are spent against the same total.
  let rowBytes = { remaining: MAX_WORKBOOK_ROW_BYTES };

  for (let sheetName of workbook.SheetNames) {
    let sheet = workbook.Sheets[sheetName];
    // Both counts are checked as the walk goes, so the sheet that crosses a ceiling is refused
    // instead of being materialized first and measured after.
    if (sheet && exceedsRowCeiling(sheet)) throw new Error(WORKBOOK_TOO_LARGE_MESSAGE);
    let rows = sheet ? readSheetRows(sheet) : [];
    rowCount += rows.length;
    if (rowCount > MAX_WORKBOOK_ROWS) throw new Error(WORKBOOK_TOO_LARGE_MESSAGE);
    let colCount = rows.reduce((widest, row) => Math.max(widest, row.length), 0);
    let sheetCells = countCells(rows);
    let chunks = chunkSheetRows(rows, rowBytes);

    cellCount += sheetCells;
    contents.push({ name: sheetName, rows, colCount, cellCount: sheetCells });
    sheets.push({ name: sheetName, chunks });
    metaSheets.push({
      name: sheetName,
      rowCount: rows.length,
      colCount,
      chunkRowStarts: chunks.map((chunk) => chunk.rowStart),
    });
  }

  // A workbook of empty sheets carries nothing to bind and nothing to summarize.
  if (cellCount === 0) throw new Error(WORKBOOK_EMPTY_MESSAGE);

  return {
    summary: renderWorkbookSummary(name, contents),
    meta: { sheets: metaSheets, cellCount },
    sheets,
  };
}

/**
 * The parser in this isolate, in the streaming shape the isolated one returns. For callers that
 * already run somewhere a runaway parse can only take itself down -- tests, and the parser worker
 * -- and as the drop-in at the workspace's call sites should the isolate ever have to be bypassed.
 */
export async function parseWorkbookInProcess(
  bytes: Uint8Array,
  mimeType: string,
  name: string,
): Promise<StreamedWorkbook> {
  let parsed = parseWorkbookAttachment(bytes, mimeType, name);
  return {
    summary: parsed.summary,
    meta: parsed.meta,
    rows: (async function* () { yield* workbookRowChunks(parsed.sheets); })(),
  };
}

// Spreadsheet attachments, parsed in the isolate instead of serialized into the prompt.
//
// Every other document type reaches the model as Markdown replayed in full on every turn. A
// workbook cannot work that way: a 490 KB / 69k-cell file becomes 1.3 MB of padded Markdown, over
// 300k tokens, more than a 200k-token window holds -- and the padding still loses the things that
// matter (dates arrive as serial numbers, side-by-side tables are squashed into one).
//
// So this module splits the workbook in two. The rows go to storage as JSON chunks the agent
// reads through a binding, and only a small deterministic Markdown *summary* -- sheet names,
// dimensions, detected table ranges, a few preview rows -- goes into the prompt. The summary is
// what the model sees every turn, so it is budgeted; the chunks are what it computes over.
//
// Parsing is pure and synchronous: no storage, no Durable Object, no network. SheetJS runs in this
// isolate, which is why the read options below switch off every part of parsing whose output
// nothing here reads.

import { Unzip, UnzipInflate } from "fflate";
import * as XLSX from "xlsx";
import { validateBindingName } from "@gadgets/workshop-shared/api";

export const XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
export const XLS_MIME_TYPE = "application/vnd.ms-excel";
export const XLSM_MIME_TYPE = "application/vnd.ms-excel.sheet.macroenabled.12";
export const XLSB_MIME_TYPE = "application/vnd.ms-excel.sheet.binary.macroenabled.12";
export const ODS_MIME_TYPE = "application/vnd.oasis.opendocument.spreadsheet";

/**
 * The spreadsheet formats that take the workbook path instead of Markdown conversion. This is the
 * spreadsheet subset of the convertible document types; the remaining convertible types (PDF,
 * DOCX, ODT) have no rows to bind and keep their existing path.
 */
export const SPREADSHEET_MIME_TYPES: ReadonlySet<string> = new Set([
  XLSX_MIME_TYPE,
  XLS_MIME_TYPE,
  XLSM_MIME_TYPE,
  XLSB_MIME_TYPE,
  ODS_MIME_TYPE,
]);

/**
 * Ceiling on the content a ZIP-based workbook inflates to, measured by inflating it.
 *
 * A Workers isolate has 128 MB; SheetJS holds roughly 130 bytes per cell with the lean read
 * options below, and ZIP-based spreadsheet formats run about 60 bytes of uncompressed XML per
 * cell -- so 32 MiB of archive content is around 500k cells and around 70 MB of heap at peak,
 * leaving room for the chunk encoding that follows. The raw upload cap does not bound this on its
 * own: a ZIP of repetitive XML compresses an order of magnitude, so a 10 MiB upload can hold far
 * more than 10 MiB of cells.
 */
export const MAX_WORKBOOK_UNCOMPRESSED_BYTES = 32 * 1024 * 1024;

/**
 * Ceiling on the rows one workbook materializes, counted across all of its sheets.
 *
 * The archive's size does not bound its shape: a sheet's stored extent runs to Excel's own
 * maximum of 1,048,576 rows, and a single cell left at the bottom of the grid -- an ordinary
 * accident, not an attack -- would otherwise cost a row array per claimed row. Anything near
 * Excel's maximum is not a dataset this binding is for; a quarter of it is already far past every
 * spreadsheet a chat has reason to carry.
 */
export const MAX_WORKBOOK_ROWS = 250_000;

/**
 * Ceiling on the JSON of all a workbook's rows: what the chunks below hold and what durable
 * storage keeps.
 *
 * Neither the archive cap nor the row ceiling bounds this on their own. A spreadsheet stores a
 * repeated string once and references it from every cell that shows it, so a few hundred rows of
 * legal, unremarkable XML can encode to hundreds of megabytes of rows -- the one attachment
 * payload with no other total-size cap in front of it.
 */
export const MAX_WORKBOOK_ROW_BYTES = 32 * 1024 * 1024;

/**
 * Ceiling on one stored row chunk. Chunks become Durable Object records, whose key and value
 * together cannot exceed 2 MB; 1 MiB leaves the envelope ample room and keeps a single `getRows`
 * page cheap to read.
 */
export const MAX_WORKBOOK_CHUNK_BYTES = 1024 * 1024;

/**
 * Ceiling on the summary, which is replayed into every model request until compaction passes the
 * message that carries it. Everything the summary elides is still reachable through the binding,
 * so truncating here costs the model nothing it cannot fetch.
 */
export const MAX_WORKBOOK_SUMMARY_BYTES = 12 * 1024;

const WORKBOOK_TOO_LARGE_MESSAGE =
  "This workbook is too large to read. Split it or remove unused sheets.";

const WORKBOOK_UNREADABLE_MESSAGE = "This spreadsheet could not be read.";

// Deliberately the same wording the document-conversion path uses for a document that yields no
// text: from the user's side a workbook with no cells and a PDF with no text are the same failure.
const WORKBOOK_EMPTY_MESSAGE = "No readable text could be extracted from this document.";

/** A cell as the summary and the stored chunks represent it. Dates arrive as ISO strings. */
export type CellValue = string | number | boolean | null;

/** One stored page of a sheet: the JSON encoding of `CellValue[][]`, starting at `rowStart`. */
export type WorkbookChunk = {
  /** 0-based index, within the sheet, of this chunk's first row. */
  rowStart: number;
  /** UTF-8 JSON array of rows. */
  bytes: Uint8Array;
};

/** Workbook shape, small enough to live alongside the attachment record. */
export type WorkbookMeta = {
  sheets: {
    name: string;
    rowCount: number;
    colCount: number;
    /** `rowStart` of each of the sheet's chunks, in order -- the index `getRows` pages against. */
    chunkRowStarts: number[];
  }[];
  cellCount: number;
};

/** Everything one parse produces: prompt text, index, and the bytes to store. */
export type ParsedWorkbook = {
  summary: string;
  meta: WorkbookMeta;
  sheets: { name: string; chunks: WorkbookChunk[] }[];
};

/** One sheet's parsed content, as the summary renderer consumes it. */
export type SheetContent = {
  name: string;
  /** Rows with trailing empty cells and trailing empty rows removed. */
  rows: CellValue[][];
  colCount: number;
  /** Non-empty cells in this sheet. */
  cellCount: number;
};

/**
 * A run of adjacent non-empty cells on a sheet, bounded by blank rows and blank columns. A `table`
 * block has a header row and data rows under it; a `text` block is a single column of loose values
 * -- a title, a note -- with nothing to pair them against.
 */
export type SheetBlock = {
  kind: "table" | "text";
  /** A1 range within the sheet, e.g. `A3:D8`; a single cell renders as `A1`. */
  range: string;
  /** Rows the block spans, including any past the ones kept below. */
  rowCount: number;
  /**
   * The block's leading rows, every row padded to the block's column span -- at most as many as
   * the summary can quote, because the rest would be copied only to be discarded. `rowCount` is
   * the block's real height.
   */
  rows: CellValue[][];
};

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

// Summary budget knobs. Past these counts the summary stops describing contents and only names
// what exists, because the binding can reach the rest and the prompt cannot afford it.
const MAX_SHEETS_WITH_BLOCKS = 100;
const MAX_BLOCKS_WITH_PREVIEW = 30;
const MAX_BLOCK_LINE_CHARS = 2000;
const PREVIEW_ROWS = 5;
const FULL_PREVIEW_ROWS = 20;
// The most rows any block can quote: a table shows its header plus a whole short block, and a
// block longer than that shows PREVIEW_ROWS of it. Nothing past this is ever rendered, so nothing
// past this is kept.
const MAX_BLOCK_KEPT_ROWS = FULL_PREVIEW_ROWS + 1;

const SUMMARY_TRUNCATION_NOTICE = "… summary truncated; the full workbook is still readable.";

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

function isBlankRow(row: CellValue[]): boolean {
  return row.every((cell) => cell === null);
}

function encodeBlockRange(
  rowStart: number,
  colStart: number,
  rowEnd: number,
  colEnd: number,
): string {
  let start = { r: rowStart, c: colStart };
  if (rowStart === rowEnd && colStart === colEnd) return XLSX.utils.encode_cell(start);
  return XLSX.utils.encode_range(start, { r: rowEnd, c: colEnd });
}

function cellAt(rows: CellValue[][], row: number, col: number): CellValue {
  return rows[row][col] ?? null;
}

/**
 * Build one block from a column run, trimming it to the rows that actually carry data in those
 * columns. Side-by-side tables share a row band but rarely the same height, so the run's own first
 * and last populated rows are what bound it -- not the band's.
 *
 * Only the rows the summary can quote are copied. A block's height is reported in full, but
 * copying a 200,000-row table to print five of its rows is work the summary never uses.
 */
function buildBlock(
  rows: CellValue[][],
  rowStart: number,
  rowEnd: number,
  colStart: number,
  colEnd: number,
): SheetBlock | null {
  let first = -1;
  let last = -1;
  for (let row = rowStart; row <= rowEnd; row++) {
    let populated = false;
    for (let col = colStart; col <= colEnd && !populated; col++) {
      populated = cellAt(rows, row, col) !== null;
    }
    if (!populated) continue;
    if (first < 0) first = row;
    last = row;
  }
  if (first < 0) return null;

  let blockRows: CellValue[][] = [];
  let keptEnd = Math.min(last, first + MAX_BLOCK_KEPT_ROWS - 1);
  for (let row = first; row <= keptEnd; row++) {
    let blockRow: CellValue[] = [];
    for (let col = colStart; col <= colEnd; col++) blockRow.push(cellAt(rows, row, col));
    blockRows.push(blockRow);
  }

  return {
    // A one-column run has no header to pair its values against, so it reads as loose text rather
    // than a table. This also covers the lone-cell case -- a sheet title, a stray note -- because
    // every column of a run holds data by construction, so a run wider than one column always
    // holds at least two cells.
    kind: colStart === colEnd ? "text" : "table",
    range: encodeBlockRange(first, colStart, last, colEnd),
    rowCount: last - first + 1,
    rows: blockRows,
  };
}

/** Split one band of consecutive non-blank rows into blocks on its fully blank columns. */
function splitBandColumns(rows: CellValue[][], rowStart: number, rowEnd: number): SheetBlock[] {
  let width = 0;
  for (let row = rowStart; row <= rowEnd; row++) width = Math.max(width, rows[row].length);

  let columnHasData: boolean[] = [];
  for (let col = 0; col < width; col++) {
    let populated = false;
    for (let row = rowStart; row <= rowEnd && !populated; row++) {
      populated = cellAt(rows, row, col) !== null;
    }
    columnHasData.push(populated);
  }

  let blocks: SheetBlock[] = [];
  let col = 0;
  while (col < width) {
    if (!columnHasData[col]) {
      col++;
      continue;
    }
    let runEnd = col;
    while (runEnd + 1 < width && columnHasData[runEnd + 1]) runEnd++;
    let block = buildBlock(rows, rowStart, rowEnd, col, runEnd);
    if (block) blocks.push(block);
    col = runEnd + 1;
  }
  return blocks;
}

/**
 * Find the table and text blocks on a sheet.
 *
 * Spreadsheets separate unrelated things with whitespace, so whitespace is the signal: fully blank
 * rows cut the sheet into bands, and within a band fully blank columns cut it into blocks. That
 * keeps a title above a table, two tables side by side, and a footnote below them from being
 * merged into one rectangle the way a naive first-row-is-the-header reading does.
 */
export function detectTables(rows: CellValue[][]): SheetBlock[] {
  let blocks: SheetBlock[] = [];
  let row = 0;

  while (row < rows.length) {
    if (isBlankRow(rows[row])) {
      row++;
      continue;
    }
    let bandEnd = row;
    while (bandEnd + 1 < rows.length && !isBlankRow(rows[bandEnd + 1])) bandEnd++;
    blocks.push(...splitBandColumns(rows, row, bandEnd));
    row = bandEnd + 1;
  }

  return blocks;
}

// Thousands separators without Intl, so the summary reads the same whatever locale data the
// runtime ships.
function formatNumber(value: number): string {
  return String(value).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

/**
 * A count with its unit, e.g. `5,501 rows`. Exported so everything that describes a workbook's
 * shape to the model -- the summary in the message, the binding description -- counts it the same
 * way; a sheet that is "5,501 rows" in one place and "5501 rows" in the other reads like two
 * different sheets.
 */
export function formatCount(value: number, unit: string): string {
  return `${formatNumber(value)} ${unit}${value === 1 ? "" : "s"}`;
}

// Unpadded CSV: quotes go on only where they change the reading, which keeps the common case as
// short as the values themselves. A cell's own line breaks are flattened first -- the summary is
// read line by line, and a value that could add a line to it could imitate the renderer's own.
function formatCsvValue(value: CellValue): string {
  if (value === null) return "";
  let text = flattenToOneLine(String(value));
  if (!/[",]/.test(text)) return text;
  return `"${text.replace(/"/g, '""')}"`;
}

function formatCsvRow(row: CellValue[]): string {
  return row.map(formatCsvValue).join(",");
}

// Summaries stay line-structured, so anything that could carry a line break is flattened first.
function flattenToOneLine(text: string): string {
  return text.replace(/[\r\n]+/g, " ");
}

/**
 * Quote a block's rows from `start`: whole while the block is small enough to be read at a glance,
 * and otherwise a fixed preview plus a count of what was left out. Rows are rendered one at a time
 * and only while they are still being shown, so a five-row preview costs five rows of rendering
 * whatever the block's height. The elided rows are never lost -- they are in the chunks the
 * binding reads.
 */
function previewParts(
  block: SheetBlock,
  start: number,
  render: (row: CellValue[]) => string,
): string[] {
  let available = block.rowCount - start;
  let shown = available <= FULL_PREVIEW_ROWS ? available : PREVIEW_ROWS;
  let parts: string[] = [];
  for (let index = start; index < start + shown; index++) parts.push(render(block.rows[index]));
  if (shown < available) parts.push(`… ${formatCount(available - shown, "more row")}`);
  return parts;
}

// Join a block's parts onto one line, stopping before a pathologically wide sheet turns a single
// block into kilobytes. Parts are appended whole, so a line is never cut mid-value.
function joinBlockLine(label: string, parts: string[]): string {
  let line = label;
  for (let [index, part] of parts.entries()) {
    let separator = index === 0 ? ": " : " / ";
    if (line.length + separator.length + part.length > MAX_BLOCK_LINE_CHARS) {
      return `${line}${separator}…`;
    }
    line += `${separator}${part}`;
  }
  return line;
}

function renderBlock(block: SheetBlock, withPreview: boolean): string {
  if (block.kind === "text") {
    let label = `Text ${block.range}`;
    if (!withPreview) return label;
    return joinBlockLine(label, previewParts(block, 0, (row) => formatCsvValue(row[0])));
  }

  // A table's first row is its header; the rows under it are what the count and the preview are
  // about.
  let label = `Table ${block.range} (${formatCount(block.rowCount - 1, "data row")})`;
  if (!withPreview) return label;
  return joinBlockLine(label, [
    formatCsvRow(block.rows[0]),
    ...previewParts(block, 1, formatCsvRow),
  ]);
}

function renderSheetHeading(sheet: SheetContent): string {
  let name = flattenToOneLine(sheet.name);
  if (sheet.rows.length === 0) return `## Sheet "${name}" (empty)`;
  let rows = formatCount(sheet.rows.length, "row");
  let cols = formatCount(sheet.colCount, "col");
  return `## Sheet "${name}" (${rows} × ${cols})`;
}

/**
 * Hold the summary under its byte budget by dropping whole trailing lines. The heading line and
 * the first sheets survive, which is the part the model needs to know what to ask the binding for.
 */
function clampToBudget(lines: string[]): string {
  let encoder = new TextEncoder();
  let summary = lines.join("\n");
  if (encoder.encode(summary).byteLength <= MAX_WORKBOOK_SUMMARY_BYTES) return summary;

  let budget =
    MAX_WORKBOOK_SUMMARY_BYTES - encoder.encode(SUMMARY_TRUNCATION_NOTICE).byteLength - 1;
  let kept: string[] = [];
  let used = 0;
  for (let line of lines) {
    let lineBytes = encoder.encode(line).byteLength + 1; // the newline that joins it
    if (used + lineBytes > budget) break;
    kept.push(line);
    used += lineBytes;
  }
  kept.push(SUMMARY_TRUNCATION_NOTICE);
  return kept.join("\n");
}

/**
 * Render the Markdown the model sees for a workbook: what the file holds and where, not what is
 * in it. Deterministic -- the same bytes always produce the same text -- because this string is
 * replayed verbatim into every turn and a drifting prompt would invalidate cached context.
 *
 * Deliberately says nothing about how the data is reached; the binding sentence is appended at
 * replay, where the binding's name is known.
 */
export function renderWorkbookSummary(name: string, sheets: SheetContent[]): string {
  let cellCount = sheets.reduce((total, sheet) => total + sheet.cellCount, 0);
  let lines = [
    `Workbook ${flattenToOneLine(name)} — ${formatCount(sheets.length, "sheet")}, ` +
      `${formatCount(cellCount, "cell")}.`,
  ];

  for (let [index, sheet] of sheets.entries()) {
    lines.push(renderSheetHeading(sheet));
    if (sheet.rows.length === 0) continue;
    // Past this many sheets the headings alone are the budget, so the rest are named and no more.
    if (index >= MAX_SHEETS_WITH_BLOCKS) continue;

    let blocks = detectTables(sheet.rows);
    // A sheet chopped into this many blocks is a layout, not a dataset: listing where they are
    // beats quoting a few cells from each.
    let withPreview = blocks.length <= MAX_BLOCKS_WITH_PREVIEW;
    for (let block of blocks) lines.push(renderBlock(block, withPreview));
  }

  return clampToBudget(lines);
}

/**
 * Split a sheet's rows into stored pages, each one JSON array of rows.
 *
 * Rows are never split across chunks, so `rowStart` alone indexes the sheet and `getRows` can find
 * the page holding any row without opening the others. A row whose JSON exceeds the cap on its own
 * has no home under that invariant, so the workbook is refused: splitting the row would break the
 * index every reader pages against, and dropping it would put a workbook in front of the model
 * with data silently missing from it.
 *
 * `budget` is the workbook's remaining row bytes, drawn down one row at a time and shared by every
 * sheet of the same workbook: encoding is where a cheap file becomes expensive data, so that is
 * where the total is refused -- while it is being spent, not once it has been.
 */
export function chunkSheetRows(
  rows: CellValue[][],
  budget: { remaining: number } = { remaining: MAX_WORKBOOK_ROW_BYTES },
): WorkbookChunk[] {
  let encoder = new TextEncoder();
  let chunks: WorkbookChunk[] = [];
  let pending: string[] = [];
  let pendingStart = 0;
  let pendingBytes = 0;

  let flush = () => {
    if (pending.length === 0) return;
    chunks.push({ rowStart: pendingStart, bytes: encoder.encode(`[${pending.join(",")}]`) });
    pending = [];
    pendingBytes = 0;
  };

  for (let [index, row] of rows.entries()) {
    let encoded = JSON.stringify(row);
    let encodedBytes = encoder.encode(encoded).byteLength;
    // Measured against a chunk holding this row alone: its JSON plus the enclosing brackets.
    if (encodedBytes + 2 > MAX_WORKBOOK_CHUNK_BYTES) throw new Error(WORKBOOK_TOO_LARGE_MESSAGE);
    budget.remaining -= encodedBytes;
    if (budget.remaining < 0) throw new Error(WORKBOOK_TOO_LARGE_MESSAGE);
    // The enclosing brackets plus one comma for each row already pending.
    let chunkBytes = pendingBytes + encodedBytes + pending.length + 2;
    if (pending.length > 0 && chunkBytes > MAX_WORKBOOK_CHUNK_BYTES) flush();
    if (pending.length === 0) pendingStart = index;
    pending.push(encoded);
    pendingBytes += encodedBytes;
  }

  flush();
  return chunks;
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
 * The env name a spreadsheet attachment is bound under in the chat's executeCode environment,
 * derived from its file name and the names already in scope.
 *
 * Every turn replays the chat log and rebuilds the binding map from scratch, so this derivation
 * must depend on nothing but the file name and the names taken before it, in message order: a name
 * that drifted between turns would leave the conversation's earlier `env.NAME` references pointing
 * at nothing, or -- worse -- at a different workbook.
 *
 * The file name's extension is deliberately kept (`big.xlsx` becomes `big_xlsx`): two exports of
 * the same report differing only in format are common, and dropping the extension would collide
 * them into `big` and `big_2`, where neither name says which file it is.
 */
export function workbookBindingName(
  fileName: string | undefined,
  taken: ReadonlySet<string>,
): string {
  // Runs of anything a JavaScript identifier cannot hold -- spaces, dots, accented letters --
  // collapse to one underscore, so the name stays readable instead of growing one underscore per
  // byte of a non-Latin file name.
  let sanitized = (fileName ?? "").replace(/[^A-Za-z0-9_]+/g, "_").replace(/^_+|_+$/g, "");
  let base: string;
  if (sanitized === "") {
    // Nothing survived sanitizing (an unnamed attachment, or a name of punctuation only), so there
    // is nothing to prefix -- the prefix is the whole name.
    base = "file";
  } else if (/^[0-9]/.test(sanitized) || !isUsableBindingName(sanitized)) {
    // A leading digit, a reserved word (`class`), or a name that would collide with an
    // Object.prototype member (`constructor`) cannot be a binding name on its own. Prefixing keeps
    // the file recognizable where dropping the offending part would not; the result is always a
    // valid name, being `file_` followed by identifier characters.
    base = `file_${sanitized}`;
  } else {
    base = sanitized;
  }

  let name = base;
  for (let suffix = 2; taken.has(name); suffix++) name = `${base}_${suffix}`;
  return name;
}

function isUsableBindingName(name: string): boolean {
  try {
    validateBindingName(name);
    return true;
  } catch {
    return false;
  }
}

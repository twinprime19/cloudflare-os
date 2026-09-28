// Spreadsheet attachments, parsed into stored rows and a summary instead of serialized into the
// prompt.
//
// Every other document type reaches the model as Markdown replayed in full on every turn. A
// workbook cannot work that way: a 490 KB / 69k-cell file becomes 1.3 MB of padded Markdown, over
// 300k tokens, more than a 200k-token window holds -- and the padding still loses the things that
// matter (dates arrive as serial numbers, side-by-side tables are squashed into one).
//
// So this module splits the workbook in two. The rows go to storage as JSON chunks the agent
// reads through a binding, and only a budgeted deterministic *summary* goes into the prompt: every
// sheet's name and dimensions, and its cells as addressed rows (see workbook-grid.ts) -- the whole
// sheet when it is small, an outline of its row bands when it is not, and just the heading once the
// workbook's budget is spent. The summary is what the model sees every turn, so it is budgeted;
// the chunks are what it computes over.
//
// This module holds the shapes, the limits and the rendering; the parse itself lives in
// workbook-parse.ts, the one module that imports SheetJS, and runs in a dynamic worker of its own
// rather than in the workspace Durable Object (see workbook-parser-isolate.ts). Nothing here
// imports the parser, so the workspace can describe and store a workbook without loading it.

import {
  ADDRESSED_ROW_LEGEND, columnLabel, renderAddressedRow, renderAddressedRows,
} from "./workbook-grid";

/**
 * Ceiling on the content a ZIP-based workbook inflates to, measured by inflating it.
 *
 * A Workers isolate has 128 MB; SheetJS holds roughly 130 bytes per cell with the lean read
 * options workbook-parse.ts uses, and ZIP-based spreadsheet formats run about 60 bytes of
 * uncompressed XML per cell -- so 32 MiB of archive content is around 500k cells and around 70 MB
 * of heap at peak, leaving room for the chunk encoding that follows. The raw upload cap does not
 * bound this on its own: a ZIP of repetitive XML compresses an order of magnitude, so a 10 MiB
 * upload can hold far more than 10 MiB of cells.
 *
 * This bounds the honest workload -- a real, dense workbook -- so it parses inside one 128 MB
 * isolate. It does not bound a file whose markup encodes repetition (a repeated ODS row, a cell
 * written without its address), and the parser's own isolate is the hard stop for those.
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
 *
 * Like the inflated-size cap, this sizes the honest workload so it fits one 128 MB isolate; it is
 * checked only once SheetJS has materialized a sheet, so for a file built to expand the parser's
 * own isolate is the hard stop.
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
 * Ceiling on the whole summary -- heading, legend and every sheet -- which is replayed into every
 * model request until compaction passes the message that carries it. Everything the summary
 * elides is still reachable through the binding, so eliding here costs the model nothing it cannot
 * fetch.
 */
export const MAX_WORKBOOK_PROMPT_BYTES = 48 * 1024;

/**
 * Ceiling on one sheet shown whole. Big enough for a financial model -- a few hundred rows of
 * labels with their numeric series collapsed -- and small enough that a data table of thousands of
 * rows never qualifies: that is a dataset to compute over, not a layout to read.
 */
export const MAX_INLINE_SHEET_BYTES = 32 * 1024;

/** Ceiling on the outline of one sheet too large to show whole. */
export const MAX_SHEET_OUTLINE_BYTES = 6 * 1024;

// The parser's failures are all one of these three, and each reaches the uploader verbatim. They
// are exported so the parser isolate's caller can tell them apart from the isolate failing.

/** The parser's refusal of a workbook over a ceiling, or more than its isolate can hold. */
export const WORKBOOK_TOO_LARGE_MESSAGE =
  "This workbook is too large to read. Split it or remove unused sheets.";

/** The parser's refusal of bytes it cannot read as a spreadsheet. */
export const WORKBOOK_UNREADABLE_MESSAGE = "This spreadsheet could not be read.";

/**
 * The parser's refusal of a workbook with no cells. Deliberately the same wording the
 * document-conversion path uses for a document that yields no text: from the user's side a
 * workbook with no cells and a PDF with no text are the same failure.
 */
export const WORKBOOK_EMPTY_MESSAGE = "No readable text could be extracted from this document.";

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

/** One stored page of rows as it travels from the parser to storage, addressed by its sheet. */
export type WorkbookRowChunk = WorkbookChunk & {
  sheetIndex: number;
  /** Position within its sheet, as `meta.sheets[sheetIndex].chunkRowStarts` indexes it. */
  chunkIndex: number;
};

/**
 * A parsed workbook whose rows arrive a chunk at a time, so the reader can store each before it
 * asks for the next and never holds the whole workbook at once.
 */
export type StreamedWorkbook = {
  summary: string;
  meta: WorkbookMeta;
  /** Every chunk of every sheet, in sheet then chunk order. */
  rows: AsyncIterable<WorkbookRowChunk>;
};

/** Parses a spreadsheet upload; throws one of the user-facing messages above when it cannot. */
export type WorkbookParser =
  (bytes: Uint8Array, mimeType: string, name: string) => Promise<StreamedWorkbook>;

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
 * A band: a run of consecutive non-blank rows, bounded by blank rows or the sheet's edges. Bands
 * are not split on blank columns -- two tables side by side share a band, and each outline row
 * shows both, header beside header, with their own column addresses.
 */
export type SheetBand = {
  /** 0-based index of the band's first row. */
  rowStart: number;
  /** 0-based index of the band's last row. */
  rowEnd: number;
  /** 0-based index of the leftmost column any of the band's rows populates. */
  colStart: number;
  /** 0-based index of the rightmost column any of the band's rows populates. */
  colEnd: number;
  /**
   * The band's first rows, whole -- at most as many as an outline quotes, because the rest would
   * be copied only to be discarded. `rowEnd` gives the band's real height.
   */
  rows: CellValue[][];
};

// Past this many sheets the summary only names what exists: the headings alone are the budget.
const MAX_SHEETS_DESCRIBED = 100;

// The rows an outline quotes from the top of each band -- a header and the first data rows under
// it, enough to show what the band's columns hold.
const OUTLINE_ROWS_PER_BAND = 3;

const OUTLINE_SUFFIX = " — outline (first rows shown; readSheet for the rest)";
const NOT_SHOWN_SUFFIX = " — not shown; use readSheet";

const SUMMARY_TRUNCATION_NOTICE = "… summary truncated; the full workbook is still readable.";

function isBlankRow(row: CellValue[]): boolean {
  return row.every((cell) => cell === null);
}

/**
 * Find the bands on a sheet: runs of consecutive non-blank rows, cut wherever a fully blank row
 * separates them.
 *
 * Spreadsheets separate unrelated things with whitespace, so a blank row is the signal that a
 * title, a table and a footnote are different things. Blank columns are deliberately not a cut:
 * splitting a band into side-by-side rectangles loses the row alignment a labelled series depends
 * on (the label in C, its values in O..AV), while addressed rows keep every cell's column anyway.
 */
export function detectBands(rows: CellValue[][]): SheetBand[] {
  let bands: SheetBand[] = [];
  let row = 0;

  while (row < rows.length) {
    if (isBlankRow(rows[row])) {
      row++;
      continue;
    }
    let rowEnd = row;
    while (rowEnd + 1 < rows.length && !isBlankRow(rows[rowEnd + 1])) rowEnd++;

    let colStart = Number.POSITIVE_INFINITY;
    let colEnd = -1;
    for (let index = row; index <= rowEnd; index++) {
      let cells = rows[index];
      let first = cells.findIndex((cell) => cell !== null);
      let last = cells.length - 1;
      while (last >= 0 && cells[last] === null) last--;
      colStart = Math.min(colStart, first);
      colEnd = Math.max(colEnd, last);
    }

    bands.push({
      rowStart: row,
      rowEnd,
      colStart,
      colEnd,
      rows: rows.slice(row, Math.min(rowEnd + 1, row + OUTLINE_ROWS_PER_BAND)),
    });
    row = rowEnd + 1;
  }

  return bands;
}

// Thousands separators without Intl, so the summary reads the same whatever locale data the
// runtime ships.
/** Thousands separators, shared with the readSheet header so "20,000 rows" there and "of 20,000" here read as one sheet. */
export function formatNumber(value: number): string {
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

// Summaries stay line-structured, so anything that could carry a line break is flattened first.
function flattenToOneLine(text: string): string {
  return text.replace(/[\r\n]+/g, " ");
}

// One encoder for every byte count the summary takes: the budgets are UTF-8 bytes, which is what
// the attachment stores and what the request-size gate measures.
const utf8 = new TextEncoder();

function byteLength(text: string): number {
  return utf8.encode(text).byteLength;
}

// Row numbers are addresses and print bare, like the row numbers that start each addressed line;
// only the count is a quantity with separators.
function renderBandLine(band: SheetBand): string {
  let rows = `rows ${band.rowStart + 1}–${band.rowEnd + 1}`;
  let cols = `${columnLabel(band.colStart)}–${columnLabel(band.colEnd)}`;
  return `${rows} (${cols}), ${formatCount(band.rowEnd - band.rowStart + 1, "row")}:`;
}

/**
 * Every band with its first `rowsPerBand` rows and a count of the rest, or `null` as soon as the
 * text crosses `maxBytes` -- rendering stops there rather than finishing a text it will discard.
 */
function renderBandsWithRows(
  bands: SheetBand[],
  rowsPerBand: number,
  maxBytes: number,
): string | null {
  let lines: string[] = [];
  let used = -1; // the first line has no newline joining it to a previous one
  let push = (line: string): boolean => {
    used += byteLength(line) + 1;
    lines.push(line);
    return used <= maxBytes;
  };

  for (let band of bands) {
    if (!push(renderBandLine(band))) return null;
    let shown = Math.min(rowsPerBand, band.rows.length);
    for (let offset = 0; offset < shown; offset++) {
      // A band's rows are non-blank by construction, so every one renders to a line.
      let line = renderAddressedRow(band.rowStart + offset, band.rows[offset], { collapse: true });
      if (line !== null && !push(line)) return null;
    }
    let height = band.rowEnd - band.rowStart + 1;
    if (height > shown && !push(`… ${formatCount(height - shown, "more row")}`)) return null;
  }
  return lines.join("\n");
}

/**
 * Band lines alone, as many as fit, then a count of the bands left out -- the last resort for a
 * sheet cut into more bands than even their one-line descriptions can afford. `null` when not even
 * the first band line fits.
 */
function renderBandLines(bands: SheetBand[], maxBytes: number): string | null {
  let lines: string[] = [];
  let used = -1;
  for (let [index, band] of bands.entries()) {
    let line = renderBandLine(band);
    let lineBytes = byteLength(line) + 1;
    let left = bands.length - index - 1;
    // Room is kept for the tail that would follow this line, so the cut always has space to say
    // how much it cut.
    let tailBytes = left === 0 ? 0 : byteLength(`… ${formatCount(left, "more band")}`) + 1;
    if (used + lineBytes + tailBytes > maxBytes) {
      if (index === 0) return null;
      lines.push(`… ${formatCount(bands.length - index, "more band")}`);
      break;
    }
    lines.push(line);
    used += lineBytes;
  }
  return lines.join("\n");
}

/**
 * The outline of a sheet too large to show whole, degrading until it fits `maxBytes`: three rows
 * per band, then one, then band lines alone. Each attempt is a fresh render, which is cheap --
 * bands are few and each quotes at most three rows.
 */
function renderOutline(bands: SheetBand[], maxBytes: number): string | null {
  return (
    renderBandsWithRows(bands, OUTLINE_ROWS_PER_BAND, maxBytes) ??
    renderBandsWithRows(bands, 1, maxBytes) ??
    renderBandLines(bands, maxBytes)
  );
}

function renderSheetHeading(sheet: SheetContent): string {
  let name = flattenToOneLine(sheet.name);
  if (sheet.rows.length === 0) return `## Sheet "${name}" (empty)`;
  let rows = formatCount(sheet.rows.length, "row");
  let cols = formatCount(sheet.colCount, "col");
  return `## Sheet "${name}" (${rows} × ${cols})`;
}

// The least a sheet can cost: its heading, marked as not shown unless there is nothing to show.
function renderUndescribedSheet(sheet: SheetContent): string {
  let heading = renderSheetHeading(sheet);
  return sheet.rows.length === 0 ? heading : `${heading}${NOT_SHOWN_SUFFIX}`;
}

/**
 * Render one sheet in the richest form `available` bytes allow, heading included: the whole sheet
 * inline; an outline of its bands; or the heading alone. Rendering the inline form stops once it
 * crosses its cap, so deciding against it never costs more than the cap's worth of rendering.
 */
function renderSheet(sheet: SheetContent, index: number, available: number): string[] {
  let heading = renderSheetHeading(sheet);
  if (sheet.rows.length === 0) return [heading];
  let notShown = [renderUndescribedSheet(sheet)];
  if (index >= MAX_SHEETS_DESCRIBED) return notShown;

  // Each body budget is what remains once its heading and the two newlines around it are paid.
  let inlineBytes = Math.min(MAX_INLINE_SHEET_BYTES, available - byteLength(heading) - 2);
  let grid = renderAddressedRows(sheet.rows, 0, { collapse: true, maxBytes: inlineBytes });
  if (!grid.truncated) return [heading, grid.text];

  let outlineHeading = `${heading}${OUTLINE_SUFFIX}`;
  let outlineBytes = Math.min(MAX_SHEET_OUTLINE_BYTES, available - byteLength(outlineHeading) - 2);
  if (outlineBytes > 0) {
    let outline = renderOutline(detectBands(sheet.rows), outlineBytes);
    if (outline !== null) return [outlineHeading, outline];
  }
  return notShown;
}

/**
 * Hold the summary under its byte budget by dropping whole trailing lines. Only reachable when a
 * workbook's headings alone overrun the budget -- everything else is budgeted as it is rendered --
 * and the heading line and the first sheets survive, which is what the model needs to know what to
 * ask the binding for.
 */
function clampToBudget(lines: string[]): string {
  let summary = lines.join("\n");
  if (byteLength(summary) <= MAX_WORKBOOK_PROMPT_BYTES) return summary;

  let budget = MAX_WORKBOOK_PROMPT_BYTES - byteLength(SUMMARY_TRUNCATION_NOTICE) - 1;
  let kept: string[] = [];
  let used = 0;
  for (let line of lines) {
    let lineBytes = byteLength(line) + 1; // the newline that joins it
    if (used + lineBytes > budget) break;
    kept.push(line);
    used += lineBytes;
  }
  kept.push(SUMMARY_TRUNCATION_NOTICE);
  return kept.join("\n");
}

/**
 * Render the text the model sees for a workbook. Deterministic -- the same bytes always produce
 * the same text -- because this string is replayed verbatim into every turn and a drifting prompt
 * would invalidate cached context.
 *
 * Sheets are rendered in workbook order, each in the richest form the remaining budget allows.
 * Every later sheet's heading is reserved up front, so a large early sheet can never push a later
 * sheet's name out of the text: the name is what the model needs to ask for the rest.
 *
 * Deliberately says nothing about how the data is reached; the binding sentence is appended at
 * replay, where the binding's name is known.
 */
export function renderWorkbookSummary(name: string, sheets: SheetContent[]): string {
  let cellCount = sheets.reduce((total, sheet) => total + sheet.cellCount, 0);
  let lines = [
    `Workbook ${flattenToOneLine(name)} — ${formatCount(sheets.length, "sheet")}, ` +
      `${formatCount(cellCount, "cell")}.`,
    ADDRESSED_ROW_LEGEND,
  ];
  let used = byteLength(lines.join("\n"));

  let headingBytes = sheets.map((sheet) => byteLength(renderUndescribedSheet(sheet)) + 1);
  let reserved = headingBytes.reduce((total, bytes) => total + bytes, 0);

  for (let [index, sheet] of sheets.entries()) {
    reserved -= headingBytes[index];
    let sheetLines = renderSheet(sheet, index, MAX_WORKBOOK_PROMPT_BYTES - used - reserved);
    lines.push(...sheetLines);
    used += byteLength(sheetLines.join("\n")) + 1;
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

/** A parse's chunks in the order they are stored: sheet by sheet, chunk by chunk. */
export function* workbookRowChunks(sheets: ParsedWorkbook["sheets"]): Generator<WorkbookRowChunk> {
  for (let [sheetIndex, sheet] of sheets.entries()) {
    for (let [chunkIndex, chunk] of sheet.chunks.entries()) {
      yield { sheetIndex, chunkIndex, rowStart: chunk.rowStart, bytes: chunk.bytes };
    }
  }
}

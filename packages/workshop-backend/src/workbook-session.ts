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
import type { WorkbookMeta } from "./chat-attachment-workbook";

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
    // Sheet names are unique within a workbook (spreadsheet software enforces it), so the first
    // match is the only match.
    let sheetIndex = this.meta.sheets.findIndex(entry => entry.name === sheet);
    if (sheetIndex < 0) {
      let names = this.meta.sheets.map(entry => JSON.stringify(entry.name)).join(", ");
      throw new Error(
          `This workbook has no sheet named ${JSON.stringify(sheet)}. Its sheets are: ${names}.`);
    }
    let {rowCount, chunkRowStarts} = this.meta.sheets[sheetIndex];

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

    let rows: CellValue[][] = [];
    for (let [chunkIndex, chunkStart] of chunkRowStarts.entries()) {
      let chunkEnd = chunkRowStarts[chunkIndex + 1] ?? rowCount;
      if (chunkEnd <= from) continue;
      if (chunkStart >= to) break;
      let stored = this.host.readWorkbookChunk(this.fileId, sheetIndex, chunkIndex);
      if (stored === undefined) {
        // The index and the pages are written in one transaction and deleted in one transaction,
        // so a gap means the records were lost, not that this range is empty. Say so rather than
        // handing back a short answer the agent would treat as the data.
        throw new Error(
            `The stored data for sheet ${JSON.stringify(sheet)} is incomplete; this workbook ` +
            `cannot be read.`);
      }
      let chunkRows = JSON.parse(new TextDecoder().decode(stored)) as CellValue[][];
      let sliceStart = Math.max(from - chunkStart, 0);
      let sliceEnd = Math.min(to - chunkStart, chunkRows.length);
      for (let index = sliceStart; index < sliceEnd; index++) rows.push(chunkRows[index]);
    }
    return rows;
  }
}

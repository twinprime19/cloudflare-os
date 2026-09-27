// The workbook parser as a dynamic worker: the module workbook-parser-isolate.ts loads into a fresh
// isolate for each spreadsheet upload. scripts/build-browser-runtime.ts bundles it, SheetJS included,
// into src/generated/workbook-parser-runtime.txt.
//
// It has no bindings and no network; it reads the bytes it is handed. `parse` does the whole parse
// and returns the summary and the index at once, but hands the rows out through a cursor, one
// stored chunk per call, so the caller can write each to storage before it asks for the next.

import { RpcTarget, WorkerEntrypoint } from "cloudflare:workers";
import { workbookRowChunks } from "../src/chat-attachment-workbook";
import type { WorkbookMeta, WorkbookRowChunk } from "../src/chat-attachment-workbook";
import { parseWorkbookAttachment } from "../src/workbook-parse";

class WorkbookRowCursor extends RpcTarget {
  #chunks: Iterator<WorkbookRowChunk>;

  constructor(chunks: Iterator<WorkbookRowChunk>) {
    super();
    this.#chunks = chunks;
  }

  next(): WorkbookRowChunk | undefined {
    let step = this.#chunks.next();
    return step.done ? undefined : step.value;
  }
}

export default class WorkbookParser extends WorkerEntrypoint {
  /**
   * Parse one upload: the summary and index at once, the rows through the cursor. Throws only the
   * parser's user-facing messages, which the caller passes through verbatim.
   */
  parse(bytes: Uint8Array, mimeType: string, name: string)
      : { summary: string; meta: WorkbookMeta; rows: WorkbookRowCursor } {
    let parsed = parseWorkbookAttachment(bytes, mimeType, name);
    return {
      summary: parsed.summary,
      meta: parsed.meta,
      rows: new WorkbookRowCursor(workbookRowChunks(parsed.sheets)),
    };
  }
}

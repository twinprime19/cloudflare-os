// Workbook parsing in a dynamic worker of its own, rather than in the workspace Durable Object.
//
// The parser's guards -- inflated size, row and row-byte ceilings -- assume a file's markup bounds
// how many cells it holds. Formats that encode repetition break that: an ODS row can declare it
// repeats a million times, and a cell written without its address costs a few bytes of XML. Such
// a file clears every guard that can run before SheetJS materializes the grid, and then
// materializes more than an isolate holds. No constant closes that class without refusing
// legitimate dense workbooks, so the parse runs where running out costs nothing: a fresh isolate
// loaded through the LOADER binding for this one upload, with its own memory budget, no bindings,
// no network and a CPU limit. If the parse exhausts it, that isolate dies and the upload is
// refused; the workspace keeps its state and its other users.
//
// The rows come back a chunk per call through a cursor the parser hands out, so the caller can
// store each before asking for the next: no RPC message carries more than one chunk, and the
// workspace never holds a whole workbook's rows.

import type { RpcTarget, WorkerEntrypoint } from "cloudflare:workers";
import {
  MAX_WORKBOOK_CHUNK_BYTES,
  WORKBOOK_EMPTY_MESSAGE,
  WORKBOOK_TOO_LARGE_MESSAGE,
  WORKBOOK_UNREADABLE_MESSAGE,
} from "./chat-attachment-workbook";
import type {
  StreamedWorkbook, WorkbookMeta, WorkbookRowChunk,
} from "./chat-attachment-workbook";
import { createWorkshopLogger } from "./observability";
// Built from worker-runtimes/workbook-parser-runtime.ts by scripts/build-browser-runtime.ts.
// SheetJS is bundled into that text and nowhere else, so the workspace's own bundle never carries
// it.
import WORKBOOK_PARSER_RUNTIME from "./generated/workbook-parser-runtime.txt";

const logger = createWorkshopLogger("workshop.workbook-parser");

/**
 * CPU the parser worker may spend on one call before the runtime stops it. A dense workbook at the
 * inflated-size cap parses in a few seconds; this is the ordinary Workers ceiling, so only a parse
 * that would never finish meets it.
 */
export const WORKBOOK_PARSE_CPU_MS = 30_000;

/** The parser worker's row cursor, as the workspace calls it. */
interface WorkbookRowCursor extends RpcTarget {
  /** The next chunk in sheet then chunk order, or `undefined` once every chunk has been sent. */
  next(): WorkbookRowChunk | undefined;
}

/** The parser worker's entrypoint (worker-runtimes/workbook-parser-runtime.ts). */
interface WorkbookParserEntrypoint extends WorkerEntrypoint {
  parse(bytes: Uint8Array, mimeType: string, name: string)
      : { summary: string; meta: WorkbookMeta; rows: WorkbookRowCursor };
}

// What `parse` resolves to on this side: the summary and index as values, the cursor as a stub,
// and a dispose that releases the stub and with it the parser isolate.
type ParseReply = Awaited<ReturnType<Fetcher<WorkbookParserEntrypoint>["parse"]>>;

// The parser's own refusals. It throws nothing else on purpose, so one of these arriving over RPC
// is the parser speaking, and the uploader is told exactly that.
const USER_FACING_MESSAGES = new Set([
  WORKBOOK_TOO_LARGE_MESSAGE,
  WORKBOOK_UNREADABLE_MESSAGE,
  WORKBOOK_EMPTY_MESSAGE,
]);

// What the runtime reports when it stops an isolate that ran past a limit -- "Worker exceeded
// memory limit.", "Worker exceeded CPU time limit." -- or when the isolate died under a call and
// took the far end of the call with it. Matched loosely until the production wording has been
// observed for each; anything unmatched is treated as a fault rather than as a file too large.
const RESOURCE_EXHAUSTION = new RegExp([
  "exceeded (?:its )?(?:memory|cpu)",
  "memory limit",
  "cpu time limit",
  "network connection lost",
  "disconnected",
].join("|"), "i");

/**
 * Turn anything the parser worker threw into the error the uploader sees.
 *
 * Three outcomes: the parser's own refusals pass through verbatim; the isolate running out of
 * memory or CPU means the file is more than a parse can hold, which is the "too large" refusal;
 * anything else -- the module failing to load, a malformed reply, a bug -- is a fault, reported as
 * unreadable and logged as an error. A fault is never worded as "too large": that would tell the
 * user to shrink a file that was never the problem.
 */
function classifyParseFailure(error: unknown): Error {
  let message = error instanceof Error ? error.message : String(error);
  // Rebuilt rather than rethrown, so the uploader sees the message and nothing the RPC layer
  // attached to the error on its way across.
  if (USER_FACING_MESSAGES.has(message)) return new Error(message);
  if (RESOURCE_EXHAUSTION.test(message)) {
    logger.warn("workbook parser isolate ran out of resources", {
      event: "attachment.workbook.isolate_exhausted", error,
    });
    return new Error(WORKBOOK_TOO_LARGE_MESSAGE, { cause: error });
  }
  logger.error("workbook parser isolate failed", {
    event: "attachment.workbook.parse_failed", error,
  });
  return new Error(WORKBOOK_UNREADABLE_MESSAGE, { cause: error });
}

type ChunkAddress = Pick<WorkbookRowChunk, "sheetIndex" | "chunkIndex" | "rowStart">;

// Every chunk the index names, in the order the cursor must send them.
function expectedChunks(meta: WorkbookMeta): ChunkAddress[] {
  return meta.sheets.flatMap((sheet, sheetIndex) =>
    sheet.chunkRowStarts.map((rowStart, chunkIndex) => ({ sheetIndex, chunkIndex, rowStart })));
}

/**
 * The cursor as an async iterable, checking each chunk against the index it is stored under.
 *
 * What the parser sends is written to storage beside an index the binding pages against, so a
 * chunk out of place, missing or oversized would leave a workbook that reads wrong; each one is
 * checked before it is handed on. The parse result -- and with it the parser isolate -- is released
 * once the last chunk is read, on the first failure, or when the reader stops early, whether or
 * not it ever started.
 */
function streamRows(
  result: ParseReply,
  meta: WorkbookMeta,
): AsyncIterable<WorkbookRowChunk> {
  let expected = expectedChunks(meta);
  let position = 0;
  let released = false;
  let release = () => {
    if (released) return;
    released = true;
    result[Symbol.dispose]();
  };

  let iterator: AsyncIterator<WorkbookRowChunk> = {
    async next() {
      if (released) return { done: true, value: undefined };
      let chunk: WorkbookRowChunk | undefined;
      try {
        chunk = await result.rows.next();
        let address = expected[position];
        if (address === undefined) {
          if (chunk !== undefined) {
            throw new Error("Parser worker sent more chunks than its index names.");
          }
          release();
          return { done: true, value: undefined };
        }
        if (chunk === undefined
            || chunk.sheetIndex !== address.sheetIndex
            || chunk.chunkIndex !== address.chunkIndex
            || chunk.rowStart !== address.rowStart
            || !(chunk.bytes instanceof Uint8Array)
            || chunk.bytes.byteLength > MAX_WORKBOOK_CHUNK_BYTES) {
          throw new Error(
              `Parser worker sent a chunk its index does not name, at position ${position}.`);
        }
      } catch (error) {
        release();
        throw classifyParseFailure(error);
      }
      position++;
      return { done: false, value: { ...expected[position - 1], bytes: chunk.bytes } };
    },
    async return() {
      release();
      return { done: true, value: undefined };
    },
  };
  return { [Symbol.asyncIterator]: () => iterator };
}

/**
 * Parse a spreadsheet upload in a dynamic worker loaded for it alone.
 *
 * Resolves once the parser has the summary and the index; the rows follow through `rows`, which the
 * caller must read to the end or stop early (`return()` on its iterator) so the isolate is
 * released. Throws, and every chunk read throws, only the parser's user-facing messages.
 *
 * Each call loads a fresh isolate, so no upload inherits another's heap. A Durable Object may have
 * ten dynamic workers with calls in flight; uploads are parsed one at a time per request and code
 * execution counts separately, so a workspace stays far below that without a queue.
 */
export async function parseWorkbookIsolated(
  loader: WorkerLoader,
  bytes: Uint8Array,
  mimeType: string,
  name: string,
): Promise<StreamedWorkbook> {
  let result: ParseReply;
  try {
    let parser = loader.load({
      compatibilityDate: "2026-02-01",
      compatibilityFlags: ["disallow_importable_env"],
      mainModule: "parser.js",
      modules: { "parser.js": WORKBOOK_PARSER_RUNTIME },
      env: {},
      // The parser reads the bytes it is handed and nothing else.
      globalOutbound: null,
      limits: { cpuMs: WORKBOOK_PARSE_CPU_MS },
    }).getEntrypoint<WorkbookParserEntrypoint>();
    result = await parser.parse(bytes, mimeType, name);
  } catch (error) {
    throw classifyParseFailure(error);
  }
  return { summary: result.summary, meta: result.meta, rows: streamRows(result, result.meta) };
}

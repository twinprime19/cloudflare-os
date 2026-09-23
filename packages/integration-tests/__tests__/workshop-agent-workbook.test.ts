// A spreadsheet attachment end to end: uploaded over the real RPC, parsed in the workspace's
// Durable Object, summarised into the replayed message, read as addressed text through the
// readSheet tool, and read back row by row through the `env.<name>` binding from inside
// executeCode's dynamic isolate.
//
// The point of running it here rather than in workshop-backend's unit suite is the loopback: the
// unit tests call WorkbookSessionImpl directly, so nothing there proves that a full page -- the
// largest `getRows` allows -- survives the hop from the isolate to the Durable Object and back.
// The sheet is therefore deliberately the worst case the cap permits.

import { afterAll, beforeAll, expect, it } from "vitest";
import * as XLSX from "xlsx";
import { z } from "zod";
import { openAgentSession } from "../src/agent-session.js";
import { startTestGatekeeperHarness, type Harness } from "../src/harness.js";
import {
  scriptedChatCompletions, SCRIPTED_MODEL_CONFIG, SCRIPTED_MODEL_ID, SCRIPTED_MODEL_PROFILE,
} from "../src/mock-model.js";
import { NetworkInterceptor } from "../src/network-interceptor.js";

const XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const FILE_NAME = "big.xlsx";
/** What `workbookBindingName()` derives from FILE_NAME, and what the scripted code below uses. */
const BINDING_NAME = "big_xlsx";
const SHEET_NAME = "Data";
/** Rows including the header, chosen to equal MAX_WORKBOOK_ROWS_PER_CALL exactly. */
const ROW_COUNT = 20_000;
const LAST_ROW = ROW_COUNT - 1;

const REGIONS = ["Northern Europe", "Latin America", "Asia Pacific", "North America"];
const CHANNELS = ["wholesale distribution", "direct enterprise", "online marketplace"];
const HEADER = [
  "OrderId", "Sku", "Description", "Region", "Channel",
  "Quantity", "UnitPrice", "Discount", "OrderedOn", "ShippedOn",
];

type SourceCell = string | number | Date;

/**
 * A wide sheet of mixed values: long-ish strings, integers, decimals and dates, so the page the
 * binding returns is realistic JSON rather than a compressible run of one type.
 */
function buildRows(): SourceCell[][] {
  const rows: SourceCell[][] = [HEADER];
  for (let n = 1; n < ROW_COUNT; n++) {
    rows.push([
      100_000 + n,
      `SKU-${String(n).padStart(6, "0")}-EMEA-RETAIL`,
      `Refurbished widget assembly ${String(n).padStart(6, "0")}`,
      REGIONS[n % REGIONS.length],
      CHANNELS[n % CHANNELS.length],
      (n % 97) + 1,
      Math.round((n % 5000) * 1.37 * 100) / 100,
      (n % 11) / 100,
      // Local components on purpose: SheetJS's writer reads a Date by its local wall clock, so a
      // UTC-built date would shift the stored cell by the host's offset.
      new Date(2026, 0, 1 + (n % 365)),
      new Date(2026, 0, 4 + (n % 365)),
    ]);
  }
  return rows;
}

/** The cell values the binding is expected to hand back: a midnight date is a plain ISO date. */
function expectedCells(row: SourceCell[]): (string | number)[] {
  return row.map(value => {
    if (!(value instanceof Date)) return value;
    const month = String(value.getMonth() + 1).padStart(2, "0");
    return `${value.getFullYear()}-${month}-${String(value.getDate()).padStart(2, "0")}`;
  });
}

// `XLSX.write` is typed `any` by SheetJS; contain that in one helper. Deflate is on because the
// stored form is what crosses the upload RPC, and the 10 MiB raw cap applies to it.
function writeWorkbook(rows: SourceCell[][]): Uint8Array {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows, { cellDates: true }),
      SHEET_NAME);
  return new Uint8Array(XLSX.write(
      book, { type: "array", bookType: "xlsx", compression: true }) as ArrayBuffer);
}

const CHAT_REQUEST = z.object({
  messages: z.array(z.object({
    role: z.string(),
    content: z.unknown(),
    tool_call_id: z.string().optional(),
  })),
});

const TOOLS_REQUEST = z.object({
  tools: z.array(z.object({ function: z.object({ name: z.string() }) })),
});

/** The names of the tools one request offered the model. */
function toolNames(request: unknown): string[] {
  return TOOLS_REQUEST.parse(request).tools.map(tool => tool.function.name);
}

/** Every text part of one request's messages, as one searchable string per role. */
function textOfRole(request: unknown, role: string): string {
  return JSON.stringify(CHAT_REQUEST.parse(request).messages
      .filter(message => message.role === role)
      .map(message => message.content));
}

function toolResult(request: unknown, toolCallId: string): string {
  const message = CHAT_REQUEST.parse(request).messages
      .find(entry => entry.role === "tool" && entry.tool_call_id === toolCallId);
  if (message === undefined) throw new Error(`No tool result for "${toolCallId}"`);
  return String(message.content);
}

let harness: Harness;
const model = scriptedChatCompletions([
  { toolCall: { id: "read-sheet", name: "readSheet",
                arguments: { file: BINDING_NAME, sheet: SHEET_NAME, range: "1:3" } } },
  {
    toolCall: {
      id: "read-workbook",
      name: "executeCode",
      arguments: {
        code: "export default async function(self, env) {\n" +
            `  const rows = await env.${BINDING_NAME}.getRows(${JSON.stringify(SHEET_NAME)});\n` +
            `  console.log(rows.length, JSON.stringify(rows[0]), ` +
            `JSON.stringify(rows[${LAST_ROW}]));\n` +
            "}\n",
      },
    },
  },
  { toolCall: { id: "describe-workbook", name: "describeBinding",
                arguments: { name: BINDING_NAME } } },
  { text: "The sheet holds 20000 rows." },
]);
const network = new NetworkInterceptor({ handlers: [model.handler] });

beforeAll(async () => {
  network.install();
  harness = await startTestGatekeeperHarness({ enableGadgetExecution: true });
});

afterAll(async () => {
  try {
    await harness?.server.close();
    expect(network.getUnmockedCalls()).toEqual([]);
  } finally {
    network.uninstall();
  }
});

it("reads a full page of an attached spreadsheet through the executeCode binding", async () => {
  const rows = buildRows();
  const content = writeWorkbook(rows);
  expect(content.byteLength).toBeLessThan(10 * 1024 * 1024);

  await using session = await openAgentSession(harness.url, {
    modelId: SCRIPTED_MODEL_ID,
    userModel: { profile: SCRIPTED_MODEL_PROFILE, config: SCRIPTED_MODEL_CONFIG },
    usernamePrefix: "workbook",
    turnTimeoutMs: 90_000,
  });

  const handle = await session.uploadAttachment({
    mimeType: XLSX_MIME_TYPE, content, name: FILE_NAME,
  });
  const turn = await session.runTurn("How many orders are in the sheet?",
      { attachments: [handle] });
  expect(turn.outcome).toEqual({ status: "completed" });
  expect(model.requests).toHaveLength(4);

  // The model never receives the rows: the message carries the structural summary plus the one
  // sentence naming the binding that reaches everything the summary elided.
  const replayedUserText = textOfRole(model.requests[0], "user");
  expect(replayedUserText).toContain(
      `Full data: readSheet(file: \\"${BINDING_NAME}\\", sheet, range) returns any range`);
  expect(replayedUserText).toContain(`in executeCode, env.${BINDING_NAME}.getRows(`);
  // A chat holding a spreadsheet is offered the tool that reads it.
  expect(toolNames(model.requests[0])).toContain("readSheet");
  expect(replayedUserText).toContain(`(${formatCount(ROW_COUNT)} rows × 10 cols)`);
  expect(replayedUserText).not.toContain(String(rows[LAST_ROW][1]));

  const sheetText = toolResult(model.requests[1], "read-sheet");
  expect(sheetText.startsWith(
      `Sheet "${SHEET_NAME}" rows 1–3 of ${formatCount(ROW_COUNT)}, columns A–J`)).toBe(true);
  expect(sheetText).toContain("\n1 A=OrderId B=Sku");
  expect(sheetText).toContain("\n2 A=100001 ");
  expect(sheetText).toContain("\n3 A=100002 ");

  const page = toolResult(model.requests[2], "read-workbook");
  expect(page).toContain(String(ROW_COUNT));
  expect(page).toContain(JSON.stringify(expectedCells(rows[0])));
  expect(page).toContain(JSON.stringify(expectedCells(rows[LAST_ROW])));

  const description = toolResult(model.requests[3], "describe-workbook");
  expect(description).toContain(`Binding: ${BINDING_NAME}`);
  expect(description).toContain(
      `${JSON.stringify(SHEET_NAME)} (${formatCount(ROW_COUNT)} rows × 10 cols)`);
  expect(description).toContain("getRows(sheet: string");

  expect(turn.history).toEqual(expect.arrayContaining([
    expect.objectContaining({
      type: "message",
      author: expect.objectContaining({ type: "agent" }),
      toolCalls: expect.arrayContaining([
        expect.objectContaining({
          toolName: "executeCode",
          output: expect.stringContaining(String(ROW_COUNT)),
        }),
      ]),
    }),
    expect.objectContaining({
      type: "message",
      toolCalls: expect.arrayContaining([
        expect.objectContaining({
          toolName: "readSheet",
          input: { file: BINDING_NAME, sheet: SHEET_NAME, range: "1:3" },
          output: sheetText,
        }),
      ]),
    }),
  ]));
  expect(model.remainingSteps()).toBe(0);
});

/** The thousands-separated form `formatCount()` renders counts in. */
function formatCount(value: number): string {
  return String(value).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

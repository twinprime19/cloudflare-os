import { z } from "zod";
import { afterAll, beforeAll, expect, it } from "vitest";
import { openAgentSession, type WorkshopAgentSession } from "../src/agent-session.js";
import {
  startTestGatekeeperHarness, TEST_VENDOR_ID, type Harness,
} from "../src/harness.js";
import {
  scriptedChatCompletions, SCRIPTED_MODEL_CONFIG, SCRIPTED_MODEL_ID,
  SCRIPTED_MODEL_PROFILE, type ChatCompletionStep, type ScriptedChatCompletions,
} from "../src/mock-model.js";
import { NetworkInterceptor } from "../src/network-interceptor.js";
import { waitFor } from "../src/rpc-client.js";

const CHAT_REQUEST = z.object({
  messages: z.array(z.object({
    role: z.string(),
    content: z.string().nullish(),
    tool_call_id: z.string().optional(),
    tool_calls: z.array(z.object({
      id: z.string(),
      type: z.literal("function"),
      function: z.object({ name: z.string(), arguments: z.string() }),
    })).optional(),
  })),
  tools: z.array(z.object({
    type: z.literal("function"),
    function: z.object({
      name: z.string(),
      description: z.string(),
      parameters: z.object({
        type: z.literal("object"),
        properties: z.record(z.string(), z.object({
          type: z.string().optional(),
          description: z.string().optional(),
        })),
        required: z.array(z.string()).optional(),
      }),
    }),
  })).optional(),
});

let harness: Harness;
const READ_TEST_VALUE =
    "export default async function(self, env) { console.log(await env.TEST_AMBIENT.readValue()); }";
const model = scriptedChatCompletions([
  {
    toolCall: {
      id: "read-test-value",
      name: "executeCode",
      arguments: {
        code: READ_TEST_VALUE,
      },
    },
  },
  { text: "The test value is 42." },
  {
    toolCall: {
      id: "read-test-value-again",
      name: "executeCode",
      arguments: { code: READ_TEST_VALUE },
    },
  },
  { text: "The test value is still 42." },
  { error: { status: 503, message: "scripted provider outage" } },
  { pending: true },
]);
// Every test shares one model endpoint; the tests that need their own script swap it in here.
let activeModel: ScriptedChatCompletions = model;
const network = new NetworkInterceptor({
  handlers: [(...args) => activeModel.handler(...args)],
});

function useScriptedModel(script: readonly ChatCompletionStep[]): ScriptedChatCompletions {
  activeModel = scriptedChatCompletions(script);
  return activeModel;
}

// Sessions stay open until the suite ends. Closing one deletes its workspace, and the abort that
// deletion lands once the session's clients are gone can crash the local workerd a moment later
// (see settleRestart) -- failing whichever sibling test is mid-flight, not the one that closed.
const sessions: WorkshopAgentSession[] = [];

// The scripted model is offered to the deployment through its one admin account, which can sign
// up only once per harness: the first session offers it, later sessions find it already shared.
let scriptedModelOffered = false;
async function openScriptedSession(): Promise<WorkshopAgentSession> {
  const userModel = scriptedModelOffered
    ? undefined
    : { profile: SCRIPTED_MODEL_PROFILE, config: SCRIPTED_MODEL_CONFIG };
  scriptedModelOffered = true;
  const session = await openAgentSession(harness.url, {
    modelId: SCRIPTED_MODEL_ID,
    userModel,
    ambientVendorIds: [TEST_VENDOR_ID],
  });
  sessions.push(session);
  return session;
}

function executeCodeStep(id: string, code: string): ChatCompletionStep {
  return { toolCall: { id, name: "executeCode", arguments: { code } } };
}

beforeAll(async () => {
  network.install();
  harness = await startTestGatekeeperHarness({ enableGadgetExecution: true });
});

afterAll(async () => {
  try {
    for (const session of sessions) await session.close();
    await harness?.server.close();
    expect(network.getUnmockedCalls()).toEqual([]);
  } finally {
    network.uninstall();
  }
});

it("keeps multi-turn history and returns provider errors", async () => {
  const session = await openScriptedSession();
  const requestsBefore = model.requests.length;
  const preCancelled = new AbortController();
  preCancelled.abort();
  const cancelledBeforeDispatch = await session.runTurn("Do not send this prompt.", {
    signal: preCancelled.signal,
  });
  expect(cancelledBeforeDispatch.outcome).toEqual({
    status: "cancelled",
    message: "Agent activation was cancelled",
  });
  expect(model.requests).toHaveLength(requestsBefore);

  const result = await session.runTurn("Read the test value and tell me what it is.");

  expect(result.outcome).toEqual({ status: "completed" });
  expect(model.requests).toHaveLength(2);
  const firstRequest = CHAT_REQUEST.parse(model.requests[0]);
  expect(firstRequest.messages).toContainEqual(expect.objectContaining({
    role: "user",
    content: expect.stringContaining("Read the test value"),
  }));
  expect(firstRequest.tools).toContainEqual(expect.objectContaining({
    type: "function",
    function: expect.objectContaining({
      name: "executeCode",
      description: expect.stringContaining("JavaScript"),
      parameters: {
        type: "object",
        properties: {
          code: {
            type: "string",
            description: expect.stringContaining("self-contained JavaScript module"),
          },
        },
        required: ["code"],
      },
    }),
  }));
  const secondRequest = CHAT_REQUEST.parse(model.requests[1]);
  expect(secondRequest.messages).toContainEqual(expect.objectContaining({
    role: "assistant",
    tool_calls: expect.arrayContaining([
      expect.objectContaining({
        id: "read-test-value",
        type: "function",
        function: expect.objectContaining({
          name: "executeCode",
          arguments: expect.stringContaining("TEST_AMBIENT"),
        }),
      }),
    ]),
  }));
  expect(secondRequest.messages).toContainEqual(expect.objectContaining({
    role: "tool",
    tool_call_id: "read-test-value",
    content: expect.stringContaining("42"),
  }));
  expect(result.history).toEqual(expect.arrayContaining([
    expect.objectContaining({
      type: "message",
      author: expect.objectContaining({ type: "agent" }),
      toolCalls: expect.arrayContaining([
        expect.objectContaining({ toolName: "executeCode", output: expect.stringContaining("42") }),
      ]),
    }),
    expect.objectContaining({
      type: "message",
      author: expect.objectContaining({ type: "agent" }),
      message: "The test value is 42.",
    }),
  ]));
  expect((await session.listActions({ filter: "observation" })).entries).toContainEqual(
      expect.objectContaining({
        type: "observation",
        description: expect.objectContaining({ title: "Read the test value" }),
      }));
  const cancelledAfterTurn = new AbortController();
  cancelledAfterTurn.abort();
  const retained = await session.runTurn("Do not send this repeated prompt.", {
    signal: cancelledAfterTurn.signal,
  });
  expect(retained.outcome.status).toBe("cancelled");
  expect(retained.history).toEqual(result.history);
  expect(retained.workpieces).toEqual(result.workpieces);
  expect(retained.usage).toEqual(result.usage);
  expect(model.requests).toHaveLength(2);
  const second = await session.runTurn("Check the test value again.");
  expect(second.outcome).toEqual({ status: "completed" });
  expect(model.requests).toHaveLength(4);
  expect(second.history).toEqual(expect.arrayContaining([
    expect.objectContaining({
      type: "message",
      author: expect.objectContaining({ type: "user" }),
      message: "Check the test value again.",
    }),
    expect.objectContaining({
      type: "message",
      author: expect.objectContaining({ type: "agent" }),
      message: "The test value is still 42.",
    }),
  ]));
  const failed = await session.runTurn("Trigger the scripted provider failure.");
  expect(failed.outcome).toMatchObject({
    status: "error",
    message: expect.stringContaining("scripted provider outage"),
  });
  expect(failed.history).toEqual(expect.arrayContaining([
    expect.objectContaining({
      type: "error",
      message: expect.stringContaining("scripted provider outage"),
    }),
  ]));
  expect(model.requests).toHaveLength(5);
  const controller = new AbortController();
  const cancelling = session.runTurn("Trigger a model request that never resolves.", {
    timeoutMs: 40_000,
    signal: controller.signal,
  });
  await waitFor("the pending model request", () =>
    Promise.resolve(model.requests.length === 6 ? true : null));
  controller.abort();
  const cancelled = await cancelling;
  expect(cancelled.outcome).toMatchObject({ status: "cancelled" });
  expect(() => session.runTurn("Do not run this turn.")).toThrow("cannot continue");
  expect(model.remainingSteps()).toBe(0);
});

it("ends the turn with a notice after three identical failing tool calls", async () => {
  // A module that fails to load (the redeclared binding is a SyntaxError), so each call is a
  // failed tool call with the same input and the same error.
  const failingCode = "export default async function() { const a = 1; const a = 2; }";
  const scripted = useScriptedModel([
    executeCodeStep("failing-1", failingCode),
    executeCodeStep("failing-2", failingCode),
    executeCodeStep("failing-3", failingCode),
    { text: "Trying a different approach." },
  ]);
  const session = await openScriptedSession();

  const stopped = await session.runTurn("Run the failing code.");

  expect(stopped.outcome).toEqual({ status: "completed" });
  expect(scripted.requests).toHaveLength(3);
  const failedCalls = stopped.history.flatMap(entry =>
    entry.type === "message" && entry.author.type === "agent" ? entry.toolCalls ?? [] : []);
  expect(failedCalls).toHaveLength(3);
  for (const call of failedCalls) {
    expect(call).toMatchObject({ toolName: "executeCode", error: expect.any(String) });
  }
  const notice = "Stopped: the same executeCode call failed 3 times in a row with the same " +
      "error. Tell me how to proceed, or ask me to try another approach.";
  expect(stopped.history.at(-1)).toMatchObject({
    type: "message",
    author: { type: "agent" },
    message: notice,
  });
  expect(stopped.history.at(-1)).not.toHaveProperty("toolCalls");

  // The chat continues, and the next turn's model sees the notice as assistant text after the
  // failed calls' results.
  const resumed = await session.runTurn("Try another approach.");

  expect(resumed.outcome).toEqual({ status: "completed" });
  expect(scripted.requests).toHaveLength(4);
  const resumedRequest = CHAT_REQUEST.parse(scripted.requests[3]);
  const noticeIndex = resumedRequest.messages.findIndex(message =>
    message.role === "assistant" && message.content === notice);
  expect(noticeIndex).toBeGreaterThan(0);
  expect(resumedRequest.messages[noticeIndex - 1]).toMatchObject({
    role: "tool",
    tool_call_id: "failing-3",
  });
  expect(resumed.history.at(-1)).toMatchObject({
    type: "message",
    author: { type: "agent" },
    message: "Trying a different approach.",
  });
  expect(scripted.remainingSteps()).toBe(0);
});

it("ends the turn with a notice at the step cap", async () => {
  const scripted = useScriptedModel(Array.from({ length: 30 }, (_, i) => executeCodeStep(
      `step-${i}`, `export default async function() { console.log(${i}); }`)));
  const session = await openScriptedSession();

  const stopped = await session.runTurn("Keep working.", { timeoutMs: 110_000 });

  expect(stopped.outcome).toEqual({ status: "completed" });
  expect(scripted.requests).toHaveLength(30);
  expect(stopped.history.at(-1)).toMatchObject({
    type: "message",
    author: { type: "agent" },
    message: "Stopped after 30 steps in one turn without finishing. Reply to let me continue.",
  });
  expect(scripted.remainingSteps()).toBe(0);
});

it("persists a visible notice when the model stops on its output cap without a reply",
    async () => {
  const scripted = useScriptedModel([
    { text: "", finishReason: "length" },
    { text: "Here is the answer." },
  ]);
  const session = await openScriptedSession();

  const stopped = await session.runTurn("Analyze the sheet.");

  expect(stopped.outcome).toEqual({ status: "completed" });
  expect(scripted.requests).toHaveLength(1);
  expect(stopped.history.at(-1)).toMatchObject({
    type: "message",
    author: { type: "agent" },
    message: "Stopped: the model used its whole output budget on reasoning and wrote no reply. " +
        "Ask a narrower question, or say \"continue\" and I will answer from where it left off.",
  });
  expect(stopped.history.at(-1)).not.toHaveProperty("toolCalls");

  const resumed = await session.runTurn("continue");

  expect(resumed.outcome).toEqual({ status: "completed" });
  expect(scripted.requests).toHaveLength(2);
  expect(resumed.history.at(-1)).toMatchObject({
    type: "message",
    author: { type: "agent" },
    message: "Here is the answer.",
  });
  expect(scripted.remainingSteps()).toBe(0);
});

it("truncates executeCode output at 32 KiB", async () => {
  const scripted = useScriptedModel([
    executeCodeStep("dump", 'export default async function() { console.log("x".repeat(40_000)); }'),
    { text: "Done." },
  ]);
  const session = await openScriptedSession();

  const result = await session.runTurn("Print a lot.");

  expect(result.outcome).toEqual({ status: "completed" });
  expect(scripted.requests).toHaveLength(2);
  const calls = result.history.flatMap(entry =>
    entry.type === "message" && entry.author.type === "agent" ? entry.toolCalls ?? [] : []);
  expect(calls).toHaveLength(1);
  const output = calls[0].toolName === "executeCode" ? calls[0].output : undefined;
  const notice = "\n… output truncated at 32 KiB (7232 bytes dropped). " +
      "Log less: compute in code and print compact results.";
  expect(output).toBe("x".repeat(32 * 1024) + notice);
  // The model received the capped output, not the full log.
  expect(CHAT_REQUEST.parse(scripted.requests[1]).messages).toContainEqual(
      expect.objectContaining({ role: "tool", tool_call_id: "dump", content: output }));
});

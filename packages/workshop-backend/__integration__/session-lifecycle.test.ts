import { createExecutionContext } from "cloudflare:test";
import { env } from "cloudflare:workers";
import { newWebSocketRpcSession, type RpcStub } from "capnweb";
import type { PublicApi } from "@gadgets/workshop-shared/api";
import server from "../src/server";
import { afterEach, describe, expect, it, vi } from "vitest";

// What a Durable Object costs is decided by how long a session keeps one alive, so every /api
// socket logs an `opened` and exactly one `closed`: an unpaired `opened` is the signature of a
// session whose browser is gone. The logger writes through console (see observability.ts), so the
// events are observed by spying on console.info, as __tests__/do-retry.test.ts does.

type LogFields = Record<string, unknown>;
type ConsoleSpy = ReturnType<typeof spyOnLogs>;

function spyOnLogs() {
  return vi.spyOn(console, "info").mockImplementation(() => {});
}

function loggedEvents(logs: ConsoleSpy, event: string): LogFields[] {
  return logs.mock.calls
      .map(call => call[0] as LogFields | undefined)
      .filter((fields): fields is LogFields => fields?.event === event);
}

// The close travels through the socket, so it lands some ticks after the client lets go.
async function waitForEvent(logs: ConsoleSpy, event: string): Promise<LogFields> {
  const deadline = Date.now() + 10_000;
  for (;;) {
    const [first] = loggedEvents(logs, event);
    if (first) return first;
    if (Date.now() > deadline) throw new Error(`No ${event} was logged within 10s.`);
    await new Promise(resolve => setTimeout(resolve, 10));
  }
}

// A service-binding fetch context ends with the upgrade response, before the socket callbacks.
// Invoke the handler directly so the WebSocket session shares the test's execution context.
async function connect(query: string): Promise<RpcStub<PublicApi>> {
  const response = await server.fetch(
    new Request(`https://workshop.invalid/api${query}`, { headers: { Upgrade: "websocket" } }),
    env, createExecutionContext());

  expect(response.status).toBe(101);
  const socket = response.webSocket;
  if (!socket) throw new TypeError("Expected a WebSocket response.");

  socket.accept();
  return newWebSocketRpcSession<PublicApi>(socket);
}

describe("api session lifecycle events", () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("reports why the browser dialled and counts the frames it sent", async () => {
    const logs = spyOnLogs();
    const publicApi = await connect("?open=redial");

    expect(loggedEvents(logs, "session.opened")).toEqual([
      expect.objectContaining({ openReason: "redial" }),
    ]);

    await publicApi.ping();
    expect(loggedEvents(logs, "session.closed")).toEqual([]);

    publicApi[Symbol.dispose]();

    const closed = await waitForEvent(logs, "session.closed");
    expect(closed).toEqual(expect.objectContaining({
      reason: "client",
      durationMs: expect.any(Number),
      inboundFrames: expect.any(Number),
    }));
    // Frames are counted, never logged one by one; the call above is at least one of them.
    expect(closed.inboundFrames as number).toBeGreaterThan(0);
    // A socket that closes and then errors is still one closed session.
    expect(loggedEvents(logs, "session.closed")).toHaveLength(1);
  });

  it("does not pass an unrecognized open reason through", async () => {
    const logs = spyOnLogs();
    const publicApi = await connect("?open=" + encodeURIComponent("not-a-reason"));
    await publicApi.ping();

    expect(loggedEvents(logs, "session.opened")).toEqual([
      expect.objectContaining({ openReason: "unknown" }),
    ]);

    // Awaited so the close lands while this test still owns the console.
    publicApi[Symbol.dispose]();
    await waitForEvent(logs, "session.closed");
  });
});

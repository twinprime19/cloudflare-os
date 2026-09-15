import { newWebSocketRpcSession, RpcTarget, type RpcStub } from "capnweb";
import { SESSION_IDLE_CLOSE_CODE } from "@gadgets/workshop-shared/api";
import { newWorkersWebSocketRpcResponse } from "../src/server";
import { describe, expect, it } from "vitest";

// Short enough to keep the suite fast, wide enough that a CI runner several times slower than a
// dev machine cannot mistake scheduling delay for an idle session. The deployed threshold is
// SESSION_IDLE_CLOSE_MS; only the mechanism is under test here, not its value.
const IDLE_CLOSE_MS = 1000;
// Well inside the threshold, so a slow round trip still lands before the timer would fire.
const CALL_INTERVAL_MS = IDLE_CLOSE_MS / 3;

interface Child {
  read(): string;
}

interface Echo {
  ping(): string;
  child(): Child;
}

// Stands in for the Durable Object stubs a real session hands out: what matters is that ending the
// session runs their disposer, which is what releases the Overseer and User DOs behind it.
class ChildTarget extends RpcTarget implements Child {
  constructor(private onDispose: () => void) {
    super();
  }

  read(): string {
    return "child";
  }

  [Symbol.dispose](): void {
    this.onDispose();
  }
}

class EchoTarget extends RpcTarget implements Echo {
  constructor(private onChildDisposed: () => void) {
    super();
  }

  ping(): string {
    return "pong";
  }

  child(): Child {
    return new ChildTarget(this.onChildDisposed);
  }
}

const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

interface Session {
  client: RpcStub<Echo>;
  closed: Promise<{ code: number; reason: string }>;
  broken: Promise<unknown>;
  isChildReleased(): boolean;
}

// A service-binding fetch context ends with the upgrade response, before the socket callbacks.
// Invoke the handler directly so the WebSocket session shares the test's execution context.
function openSession(idleCloseMs: number): Session {
  let childReleased = false;

  const response = newWorkersWebSocketRpcResponse(
    new Request("https://workshop.invalid/api", { headers: { Upgrade: "websocket" } }),
    new EchoTarget(() => { childReleased = true; }),
    { abortSignal: new AbortController().signal, idleCloseMs },
  );

  expect(response.status).toBe(101);
  const socket = response.webSocket;
  if (!socket) throw new TypeError("Expected a WebSocket response.");

  socket.accept();
  // Registered before capnweb's own handler, the way main.tsx does it: the close code has to be
  // readable before the session tears itself down.
  const closed = new Promise<{ code: number; reason: string }>(resolve => {
    socket.addEventListener("close", event => resolve({ code: event.code, reason: event.reason }));
  });
  const client = newWebSocketRpcSession<Echo>(socket);
  const broken = new Promise<unknown>(resolve => client.onRpcBroken(resolve));
  return { client, closed, broken, isChildReleased: () => childReleased };
}

describe("idle API sessions", () => {
  it("closes a socket with no inbound frames and releases the capabilities behind it", async () => {
    const session = openSession(IDLE_CLOSE_MS);
    expect(await session.client.ping()).toBe("pong");

    // Held across the quiet period, the way a workspace tab holds its Overseer. It must still be
    // alive when the socket goes quiet, or the release asserted below would prove nothing.
    const child = await session.client.child();
    expect(await child.read()).toBe("child");
    expect(session.isChildReleased()).toBe(false);

    expect(await session.closed).toEqual({ code: SESSION_IDLE_CLOSE_CODE, reason: "idle" });
    // The session is disposed behind the close, so every capability it carried goes with it: the
    // client sees a broken stub rather than a call hanging on a dead socket, and the server has
    // already dropped the objects the session was keeping alive -- the socket dying would get
    // there eventually, but the point of this backstop is to stop paying for them now.
    expect(await session.broken).toBeInstanceOf(Error);
    expect(session.isChildReleased()).toBe(true);
    await expect(session.client.ping()).rejects.toBeInstanceOf(Error);
  });

  it("keeps a socket carrying calls open past the idle threshold", async () => {
    const session = openSession(IDLE_CLOSE_MS);
    let closedEarly = false;
    void session.closed.then(() => { closedEarly = true; });

    // Twice the idle window in total, with no gap wider than a third of it.
    const deadline = Date.now() + IDLE_CLOSE_MS * 2;
    while (Date.now() < deadline) {
      expect(await session.client.ping()).toBe("pong");
      await sleep(CALL_INTERVAL_MS);
    }

    expect(closedEarly).toBe(false);
    expect(await session.client.ping()).toBe("pong");
    session.client[Symbol.dispose]();
  });
});

import {
  createExecutionContext,
  runDurableObjectAlarm,
  runInDurableObject,
} from "cloudflare:test";
import { env, exports } from "cloudflare:workers";
import { keyString } from "@gadgets/typed-storage";
import { newWebSocketRpcSession, type RpcStub } from "capnweb";
import type { AuthenticatedApi, Overseer, PublicApi } from "@gadgets/workshop-shared/api";
import { SESSION_LEASE_KEY } from "../src/blueprint-archive";
import server from "../src/server";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

// The workspace Durable Object ends itself when no browser has reached it for a lease period, which
// is the only thing that bounds a workspace held open by a front Worker session whose browser is
// gone. These tests drive the real thing end to end: a real /api socket, a real workspace, and the
// object's own alarm.
//
// Time: `vi.setSystemTime` DOES reach `Date.now()` inside the Durable Object under
// vitest-pool-workers (the object runs in this isolate), so the lease needs no test-only override.
// Only `Date` is faked -- timers stay real, because the reap's notification cap and the sockets
// themselves run on them. One consequence to respect: a delivery still travelling when the clock
// jumps is processed against the advanced clock and therefore looks like fresh activity, so a case
// that tears a session down first waits for the teardown to be logged before it moves time.
//
// Two things here are proved by construction rather than by a case, because this harness cannot
// create a gadget with committed code: the gadget-facet stub (no code means no [restore]() to
// restore a facet through -- the retained GadgetClient capability it hangs off stands in for it),
// and the calls that must NOT renew the lease -- an approval queue's, an internally constructed
// gatekeeper client's, a binding loopback's facet call -- which would each need a live gatekeeper or
// a running gadget to stage. Those paths carry no `clientActivity`, which is the whole mechanism;
// what is exercised here is the other direction: traffic that reaches the workspace renews it, and
// traffic that does not reach it (below) does not.

const PASSWORD_HASH = new Uint8Array([1, 2, 3]);

// SESSION_LEASE_MS and AGENT_RESPONSE_DELIVERED_RETENTION_MS in src/overseer.ts, which are private
// to that module. A change there that these do not mirror shows up as a failure here.
const LEASE_MS = 10 * 60_000;
const DELIVERY_RETENTION_MS = 24 * 60 * 60 * 1000;
const PAST_THE_LEASE_MS = LEASE_MS + 60_000;

// The abort reason the object uses, which the front Worker matches to close the browser's socket
// with the idle code (see vitest.integration.config.ts, which whitelists it).
const LEASE_ABORT_REASON = "idle session lease expired";
const IDLE_CLOSE_CODE = 4001;

type LogFields = Record<string, unknown>;
type SocketClose = { code: number; reason: string };
type Session = {
  publicApi: RpcStub<PublicApi>;
  // Every close this socket saw, in order. An idle release is `{4001, "idle"}`; a generic abort
  // (the degraded mode when the idle notification never lands) is anything else.
  closes: SocketClose[];
};

function spyOnLogs() {
  return vi.spyOn(console, "info").mockImplementation(() => {});
}

function loggedEvents(logs: ReturnType<typeof spyOnLogs>, event: string): LogFields[] {
  return logs.mock.calls
      .map(call => call[0] as LogFields | undefined)
      .filter((fields): fields is LogFields => fields?.event === event);
}

// Effects travel from the object to the Worker and on to the socket, so they land some ticks after
// the alarm returns. Counted in polls, never in wall-clock time: these tests move the clock.
async function waitFor<T>(what: string, poll: () => T | undefined | null): Promise<T> {
  for (let attempt = 0; attempt < 1000; ++attempt) {
    const value = poll();
    if (value !== undefined && value !== null) return value;
    await new Promise(resolve => setTimeout(resolve, 10));
  }
  throw new Error(`Timed out waiting for ${what}.`);
}

function advanceBeyondTheLease(): void {
  vi.setSystemTime(Date.now() + PAST_THE_LEASE_MS);
}

async function connect(): Promise<Session> {
  // A service-binding fetch context ends with the upgrade response, before the socket callbacks.
  // Invoke the handler directly so the WebSocket session shares the test's execution context.
  const response = await server.fetch(new Request("https://workshop.invalid/api", {
    headers: { Upgrade: "websocket" },
  }), env, createExecutionContext());

  expect(response.status).toBe(101);
  const socket = response.webSocket;
  if (!socket) throw new TypeError("Expected a WebSocket response.");

  socket.accept();
  const closes: SocketClose[] = [];
  // Registered before the Cap'n Web session so the code is recorded even though Cap'n Web tears its
  // own session down on the same event.
  socket.addEventListener("close", event => { closes.push({code: event.code, reason: event.reason}); });
  return { publicApi: newWebSocketRpcSession<PublicApi>(socket), closes };
}

async function createAccount(publicApi: RpcStub<PublicApi>, prefix: string)
    : Promise<{ username: string; token: string; userId: string }> {
  const name = prefix + crypto.randomUUID().replaceAll("-", "");
  const token = await publicApi.createAccount(name, name, PASSWORD_HASH);
  if (token === null) throw new Error(`Failed to create ${name}.`);
  return {
    username: name,
    token,
    userId: exports.UserDurableObject.idFromName(name).toString(),
  };
}

type Account = { username: string; token: string; userId: string };

// A browser session with an account, which is what every case here starts from.
async function signIn(prefix: string): Promise<{
  session: Session;
  account: Account;
  authenticated: RpcStub<AuthenticatedApi>;
}> {
  const session = await connect();
  const account = await createAccount(session.publicApi, prefix);
  const authenticated = await session.publicApi.authenticate(account.token);
  return { session, account, authenticated };
}

// What a browser's redial looks like: a second socket for an account that already has one.
async function connectAs(account: Account): Promise<RpcStub<AuthenticatedApi>> {
  const session = await connect();
  return await session.publicApi.authenticate(account.token);
}

function overseerStub(id: string) {
  return exports.OverseerDurableObject.get(exports.OverseerDurableObject.idFromString(id));
}

function scheduledAlarm(id: string): Promise<number | null> {
  return runInDurableObject(overseerStub(id), (_instance, state) => state.storage.getAlarm());
}

// Run the workspace's alarm. A reap aborts the object from inside the handler, so the invocation
// that reaped fails -- that rejection is the success signal, as in the Durable Object reset tests.
async function runAlarm(id: string): Promise<{ ran: boolean; aborted: boolean }> {
  try {
    return { ran: await runDurableObjectAlarm(overseerStub(id)), aborted: false };
  } catch (error) {
    if (error instanceof Error && error.message.includes(LEASE_ABORT_REASON)) {
      return { ran: true, aborted: true };
    }
    throw error;
  }
}

describe("workspace client-activity lease", () => {
  beforeEach(() => {
    // Only Date: faking timers as well would stop the sockets, the RPC layer and the reap's own
    // notification cap from ever running.
    vi.useFakeTimers({ toFake: ["Date"] });
  });

  afterEach(() => {
    vi.useRealTimers();
    vi.restoreAllMocks();
  });

  // (a)
  it("ends a workspace nothing has touched for a lease, and tells the browser why", async () => {
    const logs = spyOnLogs();
    const { session, authenticated } = await signIn("idle");
    const workspace = await authenticated.newGadget();
    const metadata = await workspace.getMetadata();
    // Held across the release exactly as a live editor holds them.
    const subscription = await workspace.subscribeToMetadata(() => {});

    advanceBeyondTheLease();
    // Traffic that never reaches this workspace is not activity on it: the socket's own keepalive
    // and a read of the user's own workspace list leave the lease where it was.
    await session.publicApi.ping();
    expect(await authenticated.listGadgets()).toBeInstanceOf(Array);
    await runAlarm(metadata.id);

    const close = await waitFor("the idle close", () => session.closes[0]);
    expect(close).toEqual({ code: IDLE_CLOSE_CODE, reason: "idle" });
    expect(loggedEvents(logs, "overseer.session.lease.expired")).toEqual([
      expect.objectContaining({
        gadgetId: metadata.id, idleMs: expect.any(Number), liveInterfaces: 1,
      }),
    ]);

    // The capability is gone with the object, and nothing is left scheduled: the expiry is excluded
    // from the recompute that precedes it and the alarm is not retried (`retryAlarm: false`), so the
    // next incarnation wakes only if something really needs it.
    await expect(workspace.getMetadata()).rejects.toThrow();
    expect(await scheduledAlarm(metadata.id)).toBe(null);

    subscription[Symbol.dispose]();
  });

  // (b)
  it("renews the lease on a client call and re-arms the alarm a lease after that call", async () => {
    const logs = spyOnLogs();
    const { session, authenticated } = await signIn("renew");
    const workspace = await authenticated.newGadget();
    const metadata = await workspace.getMetadata();

    // The browser is quiet but present: one read, two minutes in, is all it takes.
    vi.setSystemTime(Date.now() + 2 * 60_000);
    expect(await workspace.getMetadata()).toEqual(expect.objectContaining({ id: metadata.id }));
    const renewedAt = Date.now();

    // The alarm is armed by the first client call of an incarnation, not by every one, so it fires
    // a lease after *that* call -- here, while this renewal is still fresh. A check that finds the
    // lease renewed has to leave the lease clock alone: restarting the window from the check would
    // put the real deadline a lease after *it*, and a workspace nobody is using would get close to
    // two.
    vi.setSystemTime(renewedAt + LEASE_MS - 2 * 60_000);
    expect(await runAlarm(metadata.id)).toEqual({ ran: true, aborted: false });
    expect(loggedEvents(logs, "overseer.session.lease.expired")).toEqual([]);
    expect(session.closes).toEqual([]);
    expect(await scheduledAlarm(metadata.id)).toBe(renewedAt + LEASE_MS);

    // And that deadline is the real one: a lease after the renewal, with nobody back, it ends.
    vi.setSystemTime(renewedAt + PAST_THE_LEASE_MS);
    await runAlarm(metadata.id);

    expect(loggedEvents(logs, "overseer.session.lease.expired")).toEqual([
      expect.objectContaining({ gadgetId: metadata.id }),
    ]);
    expect(await waitFor("the idle close", () => session.closes[0]))
        .toEqual({ code: IDLE_CLOSE_CODE, reason: "idle" });
  });

  // (c)
  it("reaps a workspace whose owner kept only a subscription or a gadget capability", async () => {
    const logs = spyOnLogs();
    const { session, authenticated } = await signIn("retained");

    // The owner's capabilities are deliberately uncounted by the collaborator-access bookkeeping
    // (#liveSessions), which is why the lease counts client activity instead of live capabilities:
    // both of these workspaces have nothing "live" on them and are still held open.
    const subscribed = await authenticated.newGadget();
    const subscribedId = (await subscribed.getMetadata()).id;
    const subscription = await subscribed.subscribeToMetadata(() => {});
    subscribed[Symbol.dispose]();

    // The gadget capability a workspace mints, kept past its parent interface. Its gadget-facet
    // stub -- what a gadget UI holds -- cannot be minted here: a gadget with no committed code has
    // no [restore]() to restore a facet through. It hangs off this capability and carries this
    // capability's provenance (connectToGadget), so what it would prove is proven here.
    const held = await authenticated.newGadget();
    const heldId = (await held.getMetadata()).id;
    const gadget = await held.createGadget("retained gadget");
    held[Symbol.dispose]();

    // Both interfaces are gone from the object's point of view before the clock moves.
    await waitFor("both workspaces to be released by the client", () =>
        loggedEvents(logs, "workspace.session.closed").length === 2 || undefined);

    advanceBeyondTheLease();
    await runAlarm(subscribedId);
    await runAlarm(heldId);

    expect(loggedEvents(logs, "overseer.session.lease.expired")).toEqual([
      expect.objectContaining({ gadgetId: subscribedId, liveInterfaces: 0 }),
      expect.objectContaining({ gadgetId: heldId, liveInterfaces: 0 }),
    ]);
    expect(await scheduledAlarm(subscribedId)).toBe(null);
    expect(await scheduledAlarm(heldId)).toBe(null);
    // Nobody was left to notify, so the socket carries no idle close; the client had already let go.
    expect(session.closes).toEqual([]);
    // The retained capability died with the workspace.
    await expect(gadget.getTitle()).rejects.toThrow();

    subscription[Symbol.dispose]();
  });

  // (d)
  it("defers to outstanding agent work and reaps once the work is gone", async () => {
    const logs = spyOnLogs();
    const { session, authenticated } = await signIn("agent");
    using workspace = await authenticated.newGadget();
    const metadata = await workspace.getMetadata();

    // A recorded call to a callable agent: outstanding agent work, without a model to run it with.
    // The reap is called directly here rather than through the alarm, because the alarm's other
    // concerns would drain this record before the reap ever saw it -- the decision under test is
    // the reap's, and this is the only way to put it in front of it.
    await runInDurableObject(overseerStub(metadata.id), instance => {
      asImpl(instance).storage.pendingAgentCalls.put({
        chatId: 1, callId: 1, methodName: "doWork", args: [], argsSummary: "()",
        initiatorUserId: "unused", initiatorModelId: null,
      });
    });

    advanceBeyondTheLease();
    await runInDurableObject(overseerStub(metadata.id), instance => asImpl(instance).reapIdleSession());

    expect(loggedEvents(logs, "overseer.session.lease.expired")).toEqual([]);
    expect(session.closes).toEqual([]);
    // The work is untouched and the object still answers.
    expect(await workspace.getMetadata()).toEqual(expect.objectContaining({ id: metadata.id }));

    // The deferral is not a latch: with the work gone and nobody back, the next check ends it.
    await runInDurableObject(overseerStub(metadata.id), instance => {
      asImpl(instance).storage.pendingAgentCalls.delete(`${keyString(1)}.${keyString(1)}`);
    });
    advanceBeyondTheLease();
    await runAlarm(metadata.id);

    expect(loggedEvents(logs, "overseer.session.lease.expired")).toEqual([
      expect.objectContaining({ gadgetId: metadata.id }),
    ]);
  });

  // (e)
  it("closes the socket only for the session's last live workspace", async () => {
    const logs = spyOnLogs();
    const { session, authenticated } = await signIn("twoup");
    const first = await authenticated.newGadget();
    const firstId = (await first.getMetadata()).id;
    using second = await authenticated.newGadget();
    const secondId = (await second.getMetadata()).id;

    // The first workspace's stub leaked (the client navigated away without disposing it), so its
    // lease is what ends it -- while the second is live on the same socket.
    advanceBeyondTheLease();
    expect(await second.getMetadata()).toEqual(expect.objectContaining({ id: secondId }));
    await runAlarm(firstId);

    const closed = await waitFor("the first workspace to be reported closed", () =>
        loggedEvents(logs, "workspace.session.closed")
            .find(fields => fields.gadgetId === firstId));
    expect(closed).toEqual(expect.objectContaining({ reason: "idle" }));
    // Exactly once, and the socket is left alone: the second workspace is still live on it.
    expect(loggedEvents(logs, "workspace.session.closed")
        .filter(fields => fields.gadgetId === firstId)).toHaveLength(1);
    expect(session.closes).toEqual([]);
    expect(await second.getMetadata()).toEqual(expect.objectContaining({ id: secondId }));
    // Nothing re-opened the released workspace.
    expect(loggedEvents(logs, "workspace.session.opened")
        .filter(fields => fields.gadgetId === firstId)).toHaveLength(1);

    // Now the second one goes too: with no workspace left, the socket carries the idle close.
    first[Symbol.dispose]();
    advanceBeyondTheLease();
    await runAlarm(secondId);

    expect(await waitFor("the idle close", () => session.closes[0]))
        .toEqual({ code: IDLE_CLOSE_CODE, reason: "idle" });
  });

  // (e) two opens of one workspace
  it("keeps the socket while a second open of the same workspace is still held", async () => {
    const logs = spyOnLogs();
    const { session, authenticated } = await signIn("reopened");

    // The open the page keeps: a dialog holds this one for as long as it is on screen.
    const retained = await authenticated.newGadget();
    const retainedId = (await retained.getMetadata()).id;
    // A second open of the SAME workspace, as a list handler makes to read or rename it, released
    // as soon as it is done. What the page holds is untouched by that release.
    const temporary = await authenticated.openGadget(retainedId);
    expect(await temporary.getMetadata()).toEqual(expect.objectContaining({ id: retainedId }));
    temporary[Symbol.dispose]();
    await waitFor("the temporary open to be released", () =>
        loggedEvents(logs, "workspace.session.closed")
            .find(fields => fields.gadgetId === retainedId));

    // A different workspace on the same socket, whose stub the client leaked, so its lease ends it.
    const leaked = await authenticated.newGadget();
    const leakedId = (await leaked.getMetadata()).id;

    advanceBeyondTheLease();
    expect(await retained.getMetadata()).toEqual(expect.objectContaining({ id: retainedId }));
    await runAlarm(leakedId);

    await waitFor("the leaked workspace to be reported closed", () =>
        loggedEvents(logs, "workspace.session.closed")
            .find(fields => fields.gadgetId === leakedId));
    // The socket is left alone: the page still has an open on the first workspace, even though an
    // earlier open of that same workspace was released.
    expect(session.closes).toEqual([]);
    expect(await retained.getMetadata()).toEqual(expect.objectContaining({ id: retainedId }));

    // That surviving open is what decides: once its own lease expires, nothing is left and the
    // socket carries the idle close.
    advanceBeyondTheLease();
    await runAlarm(retainedId);

    expect(await waitFor("the idle close", () => session.closes[0]))
        .toEqual({ code: IDLE_CLOSE_CODE, reason: "idle" });
  });

  // (e) an open that never returned
  it("tells the browser to park when the lease expires while an open is still parked", async () => {
    const logs = spyOnLogs();
    const { authenticated } = await signIn("parked");
    using workspace = await authenticated.newGadget();
    const metadata = await workspace.getMetadata();

    // A collaborator's open parked on the account-configuration dialog: the workspace asks which
    // connected accounts to use and the client never answers, which is one of the sessions the
    // lease exists to end. Staged by replacing the authorization step, where that dialog is
    // awaited -- the prompt itself needs a live gatekeeper this harness cannot create. The object
    // runs in this isolate, so the parked call is observed through this counter.
    let parkedOpens = 0;
    await runInDurableObject(overseerStub(metadata.id), instance => {
      asImpl(instance).authorizeCollaborator = () => {
        ++parkedOpens;
        return new Promise(() => {});
      };
    });

    const collaboratorSession = await connect();
    const collaboratorAccount = await createAccount(collaboratorSession.publicApi, "collab");
    const collaborator =
        await collaboratorSession.publicApi.authenticate(collaboratorAccount.token);
    // Never resolves until the workspace ends the call; the rejection is collected here so it is
    // handled whenever it arrives, along with what the socket had already seen by then.
    let closesSeenWhenOpenFailed = -1;
    const parkedOpen = (async () => await collaborator.openGadget(metadata.id))()
        .then(() => undefined, (error: unknown) => {
          closesSeenWhenOpenFailed = collaboratorSession.closes.length;
          return error;
        });
    await waitFor("the collaborator's open to park", () => parkedOpens > 0 || undefined);

    advanceBeyondTheLease();
    await runAlarm(metadata.id);

    // The parked open is notified like any live interface and then fails with the abort, and its
    // socket carries the idle close, so the browser parks instead of redialling into the same wait.
    expect(await parkedOpen).toBeInstanceOf(Error);
    expect(await waitFor("the idle close on the collaborator's socket",
        () => collaboratorSession.closes[0]))
        .toEqual({ code: IDLE_CLOSE_CODE, reason: "idle" });
    // And the close is what the browser learns first: the session closes before the failed open is
    // rethrown, so what settles the call is the closed connection rather than a delivered error. A
    // client that redials on a failed open therefore cannot beat the idle code to it.
    expect(closesSeenWhenOpenFailed).toBe(1);
    expect(String(await parkedOpen)).not.toContain(LEASE_ABORT_REASON);
    expect(loggedEvents(logs, "workspace.session.open.idle")).toEqual([
      expect.objectContaining({ gadgetId: metadata.id, durationMs: expect.any(Number) }),
    ]);
    // An open that never logged `opened` gets no `closed`: the pair that measures what a session
    // holds counts only opens that started, and the owner's is the only pair here.
    expect(loggedEvents(logs, "workspace.session.opened")
        .filter(fields => fields.gadgetId === metadata.id)).toHaveLength(1);
    expect(loggedEvents(logs, "workspace.session.closed")
        .filter(fields => fields.gadgetId === metadata.id)).toHaveLength(1);
  });

  // (e) navigation
  it("leases only the workspace a navigating client still holds", async () => {
    const logs = spyOnLogs();
    const { session, authenticated } = await signIn("navigate");

    // Navigating from one workspace to the next disposes the previous one, so only the current
    // workspace is ever leased.
    const navigatedAway: string[] = [];
    for (let visit = 0; visit < 2; ++visit) {
      const workspace = await authenticated.newGadget();
      navigatedAway.push((await workspace.getMetadata()).id);
      workspace[Symbol.dispose]();
    }
    using current = await authenticated.newGadget();
    const currentId = (await current.getMetadata()).id;

    await waitFor("both earlier workspaces to be released", () =>
        loggedEvents(logs, "workspace.session.closed").length === 2 || undefined);
    // Released by the client at navigation, never by a lease.
    expect(loggedEvents(logs, "workspace.session.closed").map(fields => fields.reason))
        .toEqual(["client", "client"]);

    advanceBeyondTheLease();
    for (const id of navigatedAway) await runAlarm(id);
    // Their alarms wake objects nobody holds; the live session is untouched by them.
    expect(session.closes).toEqual([]);
    expect(await current.getMetadata()).toEqual(expect.objectContaining({ id: currentId }));

    advanceBeyondTheLease();
    await runAlarm(currentId);
    expect(await waitFor("the idle close", () => session.closes[0]))
        .toEqual({ code: IDLE_CLOSE_CODE, reason: "idle" });
  });

  // (f)
  it("releases a workspace once when the client vanishes mid-call, and never loops", async () => {
    const logs = spyOnLogs();
    const { session, authenticated } = await signIn("unmount");
    const workspace = await authenticated.newGadget();
    const metadata = await workspace.getMetadata();

    // A call still in flight as the tab goes away, which is how a session usually ends: the socket
    // closes under it and its caller is gone before it can answer.
    const lost = workspace.listActions({ filter: "action" }).catch(() => "lost");
    session.publicApi[Symbol.dispose]();
    expect(await lost).toEqual(expect.anything());
    // Let the teardown finish before the clock moves: a delivery still in the air would otherwise
    // read the advanced clock and look like fresh activity, which is a property of the test's
    // clock, not of the lease.
    await waitFor("the workspace to be released", () =>
        loggedEvents(logs, "workspace.session.closed")[0]);

    advanceBeyondTheLease();
    await runAlarm(metadata.id);
    expect(loggedEvents(logs, "overseer.session.lease.expired")).toEqual([
      expect.objectContaining({ gadgetId: metadata.id }),
    ]);

    // The incarnation the alarm wakes next has seen no browser at all, so it arms nothing and
    // aborts nothing: one release per lease, never a loop.
    expect(await runAlarm(metadata.id)).toEqual({ ran: false, aborted: false });
    expect(loggedEvents(logs, "overseer.session.lease.expired")).toHaveLength(1);
    expect(await scheduledAlarm(metadata.id)).toBe(null);
  });

  // (g) first half
  it("skips the reap when a call lands while it is reading the kill switch", async () => {
    const logs = spyOnLogs();
    const { session, authenticated } = await signIn("race");
    using workspace = await authenticated.newGadget();
    const metadata = await workspace.getMetadata();

    advanceBeyondTheLease();
    // Started, not awaited: the call below is issued while the reap is parked on its one await, and
    // the decision that follows that await sees the renewed lease.
    const reap = runInDurableObject(overseerStub(metadata.id),
        instance => asImpl(instance).reapIdleSession());
    const metadataDuringTheReap = workspace.getMetadata();

    await reap;
    expect(await metadataDuringTheReap).toEqual(expect.objectContaining({ id: metadata.id }));
    expect(loggedEvents(logs, "overseer.session.lease.expired")).toEqual([]);
    expect(session.closes).toEqual([]);
  });

  // (g) second half and (i)
  it("refuses client calls once expiry is committed, and aborts even if the notification hangs",
      async () => {
    const logs = spyOnLogs();
    const { session, account, authenticated } = await signIn("committed");
    const created = await authenticated.newGadget();
    const metadata = await created.getMetadata();
    // Hand the workspace back to the browser session, so the only interface left is the one opened
    // below, whose notification never completes.
    created[Symbol.dispose]();
    await waitFor("the browser session to let go", () =>
        loggedEvents(logs, "workspace.session.closed").length > 0 || undefined);

    // A notification that never lands is the degraded mode this is bounded for: the object must
    // still end itself, and must refuse work in the meantime.
    const neverNotified = new Promise<void>(() => {});
    const held = await overseerStub(metadata.id).open(
        account.userId, account.username, () => neverNotified) as unknown as RpcStub<Overseer>;
    expect(await held.getMetadata()).toEqual(expect.objectContaining({ id: metadata.id }));

    advanceBeyondTheLease();
    const startedAt = performance.now();
    const alarm = runAlarm(metadata.id);

    await waitFor("the expiry to be committed", () =>
        loggedEvents(logs, "overseer.session.lease.expired")[0]);
    await expect(held.getMetadata()).rejects.toThrow(/expired/);

    await alarm;
    // Bounded by the notification cap, not by the notification.
    expect(performance.now() - startedAt).toBeLessThan(30_000);
    expect(await scheduledAlarm(metadata.id)).toBe(null);
    // The socket saw no idle close: the front Worker never received the notification, which is what
    // makes this the degraded mode -- the browser reconnects instead of parking.
    expect(session.closes.some(close => close.code === IDLE_CLOSE_CODE)).toBe(false);

    // And the workspace the browser's redial re-opens expires again one lease later, rather than
    // staying up: the cost of a lost notification is one extra wake per lease, which is what the
    // expired/opened pairs in the logs measure.
    using redialled = await connectAs(account);
    using reopened = await redialled.openGadget(metadata.id);
    expect(await reopened.getMetadata()).toEqual(expect.objectContaining({ id: metadata.id }));
    advanceBeyondTheLease();
    await runAlarm(metadata.id);

    expect(loggedEvents(logs, "overseer.session.lease.expired")).toEqual([
      expect.objectContaining({ gadgetId: metadata.id }),
      expect.objectContaining({ gadgetId: metadata.id }),
    ]);
  });

  // (h) first half
  it("leaves another concern's alarm behind when it ends the object", async () => {
    const { authenticated } = await signIn("retention");
    using workspace = await authenticated.newGadget();
    const metadata = await workspace.getMetadata();

    // A delivered external-message response is kept for a day, which is a legitimate deadline the
    // expiry must not drop.
    const deliveredAt = Date.now();
    await runInDurableObject(overseerStub(metadata.id), instance => {
      asImpl(instance).storage.gadgetResponseDeliveries.put({
        idempotencyKey: "delivered-response", chatId: 1, promptSequence: 1,
        createdAt: deliveredAt, status: "delivered", deliveredAt,
      });
    });

    advanceBeyondTheLease();
    await runAlarm(metadata.id);

    expect(await scheduledAlarm(metadata.id)).toBe(deliveredAt + DELIVERY_RETENTION_MS);
  });

  // (h) second half
  it("enforces nothing while the kill switch is off", async () => {
    const logs = spyOnLogs();
    await env.BLUEPRINTS.put(SESSION_LEASE_KEY, "off");
    try {
      const { session, authenticated } = await signIn("killswitch");
      using workspace = await authenticated.newGadget();
      const metadata = await workspace.getMetadata();

      advanceBeyondTheLease();
      const checkedAt = Date.now();
      expect(await runAlarm(metadata.id)).toEqual({ ran: true, aborted: false });

      expect(loggedEvents(logs, "overseer.session.lease.skipped")).toEqual([
        expect.objectContaining({ gadgetId: metadata.id, idleMs: expect.any(Number) }),
      ]);
      expect(loggedEvents(logs, "overseer.session.lease.expired")).toEqual([]);
      expect(session.closes).toEqual([]);
      // Still armed, one lease out: flipping the switch back on takes effect at the next check.
      const alarm = await scheduledAlarm(metadata.id);
      expect(alarm! - checkedAt).toBeGreaterThan(LEASE_MS - 10_000);
      expect(alarm! - checkedAt).toBeLessThanOrEqual(LEASE_MS);
    } finally {
      await env.BLUEPRINTS.delete(SESSION_LEASE_KEY);
    }
  });
});

// The lease state lives on the workspace's OverseerImpl, which is TypeScript-private on the Durable
// Object. Reaching it is a visibility escape for the cases the alarm cannot stage on its own:
// putting outstanding agent work in front of the reap, running the reap concurrently with a client
// call, and parking an open on the step where a collaborator's account-configuration dialog is
// awaited.
type OverseerImplForTest = {
  reapIdleSession(): Promise<void>;
  authorizeCollaborator: (...args: unknown[]) => Promise<unknown>;
  storage: {
    pendingAgentCalls: { put(record: object): void; delete(key: string): void };
    gadgetResponseDeliveries: { put(record: object): void };
  };
};

function asImpl(instance: object): OverseerImplForTest {
  return (instance as { impl: OverseerImplForTest }).impl;
}

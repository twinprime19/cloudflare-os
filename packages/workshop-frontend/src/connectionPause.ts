// Deliberate socket pause for idle tabs.
//
// An open workspace tab keeps its Durable Objects awake through the live RPC session, and idle
// tabs are the bulk of that time. Pausing drops the socket on purpose and parks `main.tsx`'s
// reconnect loop instead of letting it dial back, so the Worker session and every Durable Object
// stub behind it stay released until something resumes the tab.
//
// Invariant: while `paused` is true, nothing dials. Pausing comes from the workspace idle hook
// (`useWorkspaceIdle`) and from the server closing a session whose Durable Object lease expired;
// anything that needs the server resumes -- sending from the composer, the Paused chip, leaving
// the workspace. Resuming is deliberate on purpose: a revived Durable Object stays awake until the
// next idle threshold, so merely looking at the tab must not wake it. The connection code only
// reads the flag and waits on it. No React here: `main.tsx` manages the socket outside the
// component tree.

import { SESSION_IDLE_CLOSE_CODE } from '@gadgets/workshop-shared/api'

let paused = false
let dropSocket: (() => void) | null = null
let currentSocket: WebSocket | null = null
let waiters: Array<() => void> = []
const listeners = new Set<() => void>()
const notifyListeners = () => listeners.forEach(listener => listener())

/** Registers how to drop the live socket; `main.tsx` installs this once at startup. */
export function installDropSocketHandler(handler: () => void): void {
  dropSocket = handler
}

/**
 * Records the socket `main.tsx` has just dialed, which is the only one whose close can park the
 * connection: every earlier socket is either dead or a candidate that lost its race.
 */
export function noteSocketOpened(socket: WebSocket): void {
  currentSocket = socket
}

/**
 * Parks the connection when the server closed this session because the workspace's Durable Object
 * ended its client-activity lease. Unconditional: the server only sends that code after minutes
 * without a single client call, so nothing the tab believes about itself -- busy, visible, just
 * typed in -- can contradict it, and not dialing back is exactly the deliberate-resume rule. A
 * close arriving from a socket already replaced says nothing about the live connection, so it is
 * ignored rather than parking a healthy one. Every other close code is an ordinary outage and
 * belongs to the reconnect loop.
 */
export function noteSocketClosed(socket: WebSocket, closeCode: number): void {
  if (closeCode !== SESSION_IDLE_CLOSE_CODE || socket !== currentSocket) return
  pauseConnection()
}

/** True between `pauseConnection()` and the next `resumeConnection()`. */
export function isConnectionPaused(): boolean {
  return paused
}

/**
 * Parks the connection. Drops the live socket if there is one; during an outage only the flag is
 * set and the reconnect loop parks at its next iteration. Idempotent.
 */
export function pauseConnection(): void {
  if (paused) return
  paused = true
  dropSocket?.()
  notifyListeners()
}

/** Releases every waiter so the reconnect loop dials again at once. Idempotent. */
export function resumeConnection(): void {
  if (!paused) return
  paused = false
  const pending = waiters
  waiters = []
  pending.forEach(resolve => resolve())
  notifyListeners()
}

/** Hears every pause and resume, so UI can mirror the flag; returns the unsubscribe. */
export function subscribeConnectionPause(listener: () => void): () => void {
  listeners.add(listener)
  return () => { listeners.delete(listener) }
}

/**
 * Resolves immediately when not paused, otherwise on the next `resumeConnection()`. The result
 * says whether it actually waited, so a caller can skip a backoff it no longer owes.
 */
export function waitWhilePaused(): Promise<boolean> {
  if (!paused) return Promise.resolve(false)
  return new Promise(resolve => waiters.push(() => resolve(true)))
}

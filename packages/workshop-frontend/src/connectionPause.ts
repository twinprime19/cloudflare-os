// Deliberate socket pause for idle tabs.
//
// An open workspace tab keeps its Durable Objects awake through the live RPC session, and idle
// tabs are the bulk of that time. Pausing drops the socket on purpose and parks `main.tsx`'s
// reconnect loop instead of letting it dial back, so the Worker session and every Durable Object
// stub behind it stay released until something resumes the tab.
//
// Invariant: while `paused` is true, nothing dials. Pausing comes from the workspace idle hook
// (`useWorkspaceIdle`), or from the server's own idle close while that hook is mounted and agrees
// the tab is idle; anything that needs the server resumes -- sending from the composer, the
// Paused chip, leaving the workspace. Resuming is deliberate on purpose: a revived Durable Object
// stays awake until the next idle threshold, so merely looking at the tab must not wake it. The
// connection code only reads the flag and waits on it. No React here: `main.tsx` manages the
// socket outside the component tree.

let paused = false
let dropSocket: (() => void) | null = null
let idleClosePolicy: (() => boolean) | null = null
let waiters: Array<() => void> = []
const listeners = new Set<() => void>()
const notifyListeners = () => listeners.forEach(listener => listener())

/** Registers how to drop the live socket; `main.tsx` installs this once at startup. */
export function installDropSocketHandler(handler: () => void): void {
  dropSocket = handler
}

/**
 * Registers who decides what a server-side idle close means, or `null` to clear it. Only the
 * workspace idle hook installs one; routes without it reconnect after an idle close the way they
 * do after any other disconnect.
 */
export function installIdleClosePolicy(policy: (() => boolean) | null): void {
  idleClosePolicy = policy
}

/** True when the server closing an idle socket should park the connection instead of redialing. */
export function shouldPauseOnIdleClose(): boolean {
  return idleClosePolicy?.() ?? false
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

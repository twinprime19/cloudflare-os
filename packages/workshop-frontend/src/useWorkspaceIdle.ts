import { useEffect, useRef } from 'react'
import {
  isConnectionPaused,
  pauseConnection,
  resumeConnection,
  subscribeConnectionPause,
} from './connectionPause'
import { useConnectionPaused } from './useConnectionPaused'

/** Hidden tabs pause after this long. */
export const HIDDEN_PAUSE_MS = 5 * 60_000
/** Visible tabs pause after this long without input. */
export const VISIBLE_IDLE_PAUSE_MS = 10 * 60_000
/** Poll cadence. An interval with timestamps survives background-tab throttling; one long timeout does not. */
export const IDLE_TICK_MS = 30_000

// Input that defers a pause: the user is here and may be about to ask for something. Mousemove is
// deliberately absent: a hand resting on a trackpad or a cursor nudged by a notification would
// keep a tab awake for hours.
const ACTIVITY_EVENTS = ['pointerdown', 'keydown', 'touchstart', 'wheel', 'scroll'] as const

// Module scope, not effect scope: the sandboxed gadget iframe reports its input through
// `noteWorkspaceActivity()` rather than through this window's listeners. The effect re-stamps it on
// mount, so a stale value from a previous workspace cannot shorten the first window.
let lastActivityAt = Date.now()

/**
 * Records input this window cannot observe. The sandboxed gadget iframe forwards its own
 * `pointerdown`/`keydown`/`wheel`/`touchstart` here, so working inside a gadget defers the pause
 * exactly like working in the parent page does.
 */
export function noteWorkspaceActivity(): void {
  lastActivityAt = Date.now()
}

/**
 * Pauses the workspace's RPC connection when the tab has been hidden or without input long enough,
 * and reports whether it is paused. It never resumes on its own: reading what is already on screen
 * costs nothing, and a revived Durable Object stays awake until the next threshold, so only a
 * deliberate request (sending from the composer, answering an action or connection request, the
 * Paused chip, leaving the workspace) calls `resumeConnection`. `busy` blocks entering a pause but never forces an exit, so a threshold that
 * elapsed while an agent was streaming lands on the next tick after it finishes.
 *
 * Only real input defers the pause. A focused gadget iframe proves nothing on its own -- switching
 * apps or walking away leaves it focused -- so the iframe reports its input through
 * `noteWorkspaceActivity` instead. Unmounting while paused resumes, so the parked reconnect loop
 * never outlives the only thing that would wake it.
 */
export function useWorkspaceIdle(busy: boolean): boolean {
  const paused = useConnectionPaused()
  const busyRef = useRef(busy)
  busyRef.current = busy

  useEffect(() => {
    noteWorkspaceActivity()
    let hiddenAt: number | null = document.visibilityState === 'hidden' ? Date.now() : null

    const onActivity = noteWorkspaceActivity
    const onVisibilityChange = () => {
      if (document.visibilityState === 'hidden') {
        hiddenAt = Date.now()
        return
      }
      hiddenAt = null
      noteWorkspaceActivity()
    }
    const tick = () => {
      if (isConnectionPaused() || busyRef.current) return
      const now = Date.now()
      if (hiddenAt !== null) {
        if (now - hiddenAt >= HIDDEN_PAUSE_MS) pauseConnection()
        return
      }
      if (now - lastActivityAt >= VISIBLE_IDLE_PAUSE_MS) pauseConnection()
    }

    const unsubscribe = subscribeConnectionPause(() => {
      // A resume is a request in flight; give it a full idle window before the next pause.
      if (!isConnectionPaused()) noteWorkspaceActivity()
    })
    for (const type of ACTIVITY_EVENTS) {
      window.addEventListener(type, onActivity, { capture: true, passive: true })
    }
    document.addEventListener('visibilitychange', onVisibilityChange)
    const interval = setInterval(tick, IDLE_TICK_MS)

    return () => {
      for (const type of ACTIVITY_EVENTS) {
        window.removeEventListener(type, onActivity, { capture: true })
      }
      document.removeEventListener('visibilitychange', onVisibilityChange)
      clearInterval(interval)
      unsubscribe()
      resumeConnection()
    }
  }, [])

  return paused
}

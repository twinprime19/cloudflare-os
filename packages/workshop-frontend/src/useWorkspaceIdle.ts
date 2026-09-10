import { useEffect, useRef, useState } from 'react'
import {
  isConnectionPaused,
  pauseConnection,
  resumeConnection,
  subscribeConnectionPause,
} from './connectionPause'

/** Hidden tabs pause after this long. */
export const HIDDEN_PAUSE_MS = 5 * 60_000
/** Visible tabs pause after this long without input. */
export const VISIBLE_IDLE_PAUSE_MS = 15 * 60_000
/** Poll cadence. An interval with timestamps survives background-tab throttling; one long timeout does not. */
export const IDLE_TICK_MS = 30_000

// Input that defers a pause: the user is here and may be about to ask for something. Mousemove is
// deliberately absent: a hand resting on a trackpad or a cursor nudged by a notification would
// keep a tab awake for hours.
const ACTIVITY_EVENTS = ['pointerdown', 'keydown', 'touchstart', 'wheel', 'scroll'] as const

/**
 * Pauses the workspace's RPC connection when the tab has been hidden or without input long enough,
 * and reports whether it is paused. It never resumes on its own: reading what is already on screen
 * costs nothing, and a revived Durable Object stays awake until the next threshold, so only a
 * deliberate request (sending from the composer, answering an action or connection request, the
 * Paused chip, leaving the workspace) calls `resumeConnection`. `busy` blocks entering a pause but never forces an exit, so a threshold that
 * elapsed while an agent was streaming lands on the next tick after it finishes.
 *
 * Focus inside the gadget iframe counts as input at each tick: its events never reach this window,
 * so a user working inside the gadget would otherwise look idle. Unmounting while paused resumes,
 * so the parked reconnect loop never outlives the only thing that would wake it.
 */
export function useWorkspaceIdle(busy: boolean): boolean {
  const [paused, setPaused] = useState(isConnectionPaused)
  const busyRef = useRef(busy)
  busyRef.current = busy

  useEffect(() => {
    let lastActivityAt = Date.now()
    let hiddenAt: number | null = document.visibilityState === 'hidden' ? Date.now() : null

    const onActivity = () => { lastActivityAt = Date.now() }
    const onVisibilityChange = () => {
      if (document.visibilityState === 'hidden') {
        hiddenAt = Date.now()
        return
      }
      hiddenAt = null
      lastActivityAt = Date.now()
    }
    const tick = () => {
      if (isConnectionPaused() || busyRef.current) return
      const now = Date.now()
      if (document.activeElement?.tagName === 'IFRAME') lastActivityAt = now
      if (hiddenAt !== null) {
        if (now - hiddenAt >= HIDDEN_PAUSE_MS) pauseConnection()
        return
      }
      if (now - lastActivityAt >= VISIBLE_IDLE_PAUSE_MS) pauseConnection()
    }

    const unsubscribe = subscribeConnectionPause(() => {
      setPaused(isConnectionPaused())
      // A resume is a request in flight; give it a full idle window before the next pause.
      if (!isConnectionPaused()) lastActivityAt = Date.now()
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

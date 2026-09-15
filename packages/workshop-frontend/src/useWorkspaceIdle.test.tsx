// @vitest-environment jsdom
/* eslint-disable react/react-in-jsx-scope */

import { act } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  installDropSocketHandler,
  isConnectionPaused,
  resumeConnection,
  shouldPauseOnIdleClose,
} from './connectionPause'
import {
  HIDDEN_PAUSE_MS,
  IDLE_TICK_MS,
  VISIBLE_IDLE_PAUSE_MS,
  noteWorkspaceActivity,
  useWorkspaceIdle,
} from './useWorkspaceIdle'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

function elapse(ms: number) {
  act(() => { vi.advanceTimersByTime(ms) })
}

function fire(type: string) {
  act(() => { window.dispatchEvent(new Event(type, { bubbles: true })) })
}

describe('useWorkspaceIdle', () => {
  let root: Root | undefined
  let container: HTMLDivElement
  let latest: boolean
  let visibility: DocumentVisibilityState
  const drop = vi.fn<() => void>()

  function Probe({ busy }: { busy: boolean }) {
    latest = useWorkspaceIdle(busy)
    return null
  }

  async function mount(busy = false) {
    await act(async () => root!.render(<Probe busy={busy} />))
  }

  function setVisibility(state: DocumentVisibilityState) {
    visibility = state
    act(() => { document.dispatchEvent(new Event('visibilitychange')) })
  }

  beforeEach(() => {
    vi.useFakeTimers()
    visibility = 'visible'
    Object.defineProperty(document, 'visibilityState', { configurable: true, get: () => visibility })
    drop.mockClear()
    installDropSocketHandler(drop)
    container = document.createElement('div')
    document.body.append(container)
    root = createRoot(container)
  })

  afterEach(() => {
    act(() => root?.unmount())
    root = undefined
    container.remove()
    vi.useRealTimers()
  })

  it('pauses a hidden tab after the hidden threshold and stays paused when it is visible again', async () => {
    await mount()
    setVisibility('hidden')
    elapse(HIDDEN_PAUSE_MS - IDLE_TICK_MS)
    expect(latest).toBe(false)
    expect(isConnectionPaused()).toBe(false)

    elapse(IDLE_TICK_MS)
    expect(latest).toBe(true)
    expect(isConnectionPaused()).toBe(true)
    expect(drop).toHaveBeenCalledTimes(1)

    elapse(IDLE_TICK_MS * 4)
    expect(drop).toHaveBeenCalledTimes(1)

    // Looking at the tab again is not a request; only something that needs the server resumes.
    setVisibility('visible')
    expect(latest).toBe(true)
    expect(isConnectionPaused()).toBe(true)

    act(() => resumeConnection())
    expect(latest).toBe(false)
    expect(isConnectionPaused()).toBe(false)
  })

  it('pauses a visible tab after the idle threshold; input alone does not resume it', async () => {
    await mount()
    elapse(VISIBLE_IDLE_PAUSE_MS)
    expect(latest).toBe(true)
    expect(isConnectionPaused()).toBe(true)
    expect(drop).toHaveBeenCalledTimes(1)

    for (const type of ['mousemove', 'pointerdown', 'keydown', 'scroll']) {
      fire(type)
      expect(latest).toBe(true)
      expect(isConnectionPaused()).toBe(true)
    }
    elapse(IDLE_TICK_MS * 4)
    expect(drop).toHaveBeenCalledTimes(1)

    act(() => resumeConnection())
    expect(latest).toBe(false)
  })

  it('treats a click as input that defers the idle pause, and mousemove as nothing', async () => {
    await mount()
    elapse(VISIBLE_IDLE_PAUSE_MS / 2)
    fire('pointerdown')
    elapse(VISIBLE_IDLE_PAUSE_MS / 2)
    expect(latest).toBe(false)

    fire('mousemove')
    elapse(VISIBLE_IDLE_PAUSE_MS / 2)
    expect(latest).toBe(true)
  })

  it('never pauses while busy, then pauses on the first tick after busy clears', async () => {
    await mount(true)
    elapse(VISIBLE_IDLE_PAUSE_MS * 2)
    expect(latest).toBe(false)
    expect(isConnectionPaused()).toBe(false)

    await mount(false)
    expect(latest).toBe(false)
    elapse(IDLE_TICK_MS)
    expect(latest).toBe(true)
    expect(isConnectionPaused()).toBe(true)
  })

  it('pauses a tab whose gadget iframe is focused but untouched', async () => {
    const iframe = document.createElement('iframe')
    iframe.tabIndex = 0  // jsdom only focuses elements it deems focusable
    document.body.append(iframe)  // outside the root's container, which React clears on render
    await mount()
    iframe.focus()
    expect(document.activeElement).toBe(iframe)

    elapse(VISIBLE_IDLE_PAUSE_MS)
    expect(latest).toBe(true)
    expect(isConnectionPaused()).toBe(true)
    expect(drop).toHaveBeenCalledTimes(1)
    iframe.remove()
  })

  it('defers the pause when the iframe reports input through noteWorkspaceActivity', async () => {
    await mount()
    elapse(VISIBLE_IDLE_PAUSE_MS - IDLE_TICK_MS * 2)
    act(() => noteWorkspaceActivity())

    elapse(IDLE_TICK_MS * 4)
    expect(latest).toBe(false)
    expect(isConnectionPaused()).toBe(false)

    elapse(VISIBLE_IDLE_PAUSE_MS - IDLE_TICK_MS * 2)
    expect(latest).toBe(true)
    expect(isConnectionPaused()).toBe(true)
  })

  it('pauses again after a resume once the idle threshold elapses anew', async () => {
    await mount()
    elapse(VISIBLE_IDLE_PAUSE_MS)
    expect(latest).toBe(true)
    fire('keydown')
    act(() => resumeConnection())
    expect(latest).toBe(false)

    elapse(VISIBLE_IDLE_PAUSE_MS - IDLE_TICK_MS)
    expect(latest).toBe(false)
    elapse(IDLE_TICK_MS)
    expect(latest).toBe(true)
    expect(drop).toHaveBeenCalledTimes(2)
  })

  it('parks on a server idle close only while mounted, idle, and not busy', async () => {
    expect(shouldPauseOnIdleClose()).toBe(false)

    await mount()
    elapse(60_000)
    expect(shouldPauseOnIdleClose()).toBe(false)

    // A gadget the user works in without calling the server goes quiet on the wire while the tab
    // is anything but idle; the forwarded input is the only evidence, so it has to win here.
    act(() => noteWorkspaceActivity())
    elapse(60_000)
    expect(shouldPauseOnIdleClose()).toBe(false)

    elapse(VISIBLE_IDLE_PAUSE_MS)
    expect(shouldPauseOnIdleClose()).toBe(true)

    await mount(true)
    expect(shouldPauseOnIdleClose()).toBe(false)

    await mount(false)
    act(() => root!.unmount())
    root = undefined
    expect(shouldPauseOnIdleClose()).toBe(false)
  })

  it('keeps parking on a server idle close after the tab merely became visible', async () => {
    await mount()
    elapse(VISIBLE_IDLE_PAUSE_MS)
    expect(shouldPauseOnIdleClose()).toBe(true)

    // A thawed tab fires visibilitychange before the close event it queued while frozen; looking
    // at the tab is not input, so the close still parks it.
    setVisibility('hidden')
    setVisibility('visible')
    expect(shouldPauseOnIdleClose()).toBe(true)

    act(() => noteWorkspaceActivity())
    expect(shouldPauseOnIdleClose()).toBe(false)
  })

  it('resumes the connection when unmounted while paused', async () => {
    await mount()
    setVisibility('hidden')
    elapse(HIDDEN_PAUSE_MS)
    expect(isConnectionPaused()).toBe(true)

    act(() => root!.unmount())
    root = undefined
    expect(isConnectionPaused()).toBe(false)
  })
})

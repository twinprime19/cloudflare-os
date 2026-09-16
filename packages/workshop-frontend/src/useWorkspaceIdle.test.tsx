// @vitest-environment jsdom
/* eslint-disable react/react-in-jsx-scope */

import { act } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { installDropSocketHandler, isConnectionPaused, resumeConnection } from './connectionPause'
import { HIDDEN_PAUSE_MS, IDLE_TICK_MS, VISIBLE_IDLE_PAUSE_MS, useWorkspaceIdle } from './useWorkspaceIdle'

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

  it('counts focus inside an iframe as input', async () => {
    const iframe = document.createElement('iframe')
    iframe.tabIndex = 0  // jsdom only focuses elements it deems focusable
    document.body.append(iframe)  // outside the root's container, which React clears on render
    await mount()
    iframe.focus()
    expect(document.activeElement).toBe(iframe)

    elapse(VISIBLE_IDLE_PAUSE_MS * 2)
    expect(latest).toBe(false)
    expect(isConnectionPaused()).toBe(false)

    iframe.blur()
    elapse(VISIBLE_IDLE_PAUSE_MS)
    expect(latest).toBe(true)
    iframe.remove()
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

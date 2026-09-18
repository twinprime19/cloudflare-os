import { describe, expect, it, vi } from 'vitest'
import { SESSION_IDLE_CLOSE_CODE } from '@gadgets/workshop-shared/api'

// The module keeps its flag and waiters in module scope, the way `main.tsx` keeps the connection
// itself, so each test imports a fresh copy rather than sharing state through the suite.
async function freshModule() {
  vi.resetModules()
  return await import('./connectionPause')
}

// The pause policy compares sockets by identity only, so a bare object stands in for one.
const socket = () => ({}) as WebSocket

describe('connectionPause', () => {
  it('drops the socket once and stays paused on repeated pauses', async () => {
    const m = await freshModule()
    const drop = vi.fn<() => void>()
    m.installDropSocketHandler(drop)

    expect(m.isConnectionPaused()).toBe(false)
    m.pauseConnection()
    m.pauseConnection()

    expect(m.isConnectionPaused()).toBe(true)
    expect(drop).toHaveBeenCalledTimes(1)
  })

  it('does not throw when paused before a drop handler is installed', async () => {
    const m = await freshModule()
    expect(() => m.pauseConnection()).not.toThrow()
    expect(m.isConnectionPaused()).toBe(true)
  })

  it('resolves waitWhilePaused with false immediately when not paused', async () => {
    const m = await freshModule()
    await expect(m.waitWhilePaused()).resolves.toBe(false)
  })

  it('parks every waiter until resume, then releases all of them with true', async () => {
    const m = await freshModule()
    m.installDropSocketHandler(() => {})
    m.pauseConnection()

    let settled = 0
    const first = m.waitWhilePaused().then(waited => { settled++; return waited })
    const second = m.waitWhilePaused().then(waited => { settled++; return waited })
    await Promise.resolve()
    expect(settled).toBe(0)

    m.resumeConnection()
    await expect(first).resolves.toBe(true)
    await expect(second).resolves.toBe(true)
    expect(m.isConnectionPaused()).toBe(false)
  })

  it('parks on the server idle close of the socket in use, whatever the tab is doing', async () => {
    const m = await freshModule()
    const drop = vi.fn<() => void>()
    m.installDropSocketHandler(drop)
    const live = socket()
    m.noteSocketOpened(live)

    // Every other close code is an ordinary outage: the reconnect loop owns those.
    m.noteSocketClosed(live, 1006)
    expect(m.isConnectionPaused()).toBe(false)
    expect(drop).not.toHaveBeenCalled()

    m.noteSocketClosed(live, SESSION_IDLE_CLOSE_CODE)
    expect(m.isConnectionPaused()).toBe(true)
    expect(drop).toHaveBeenCalledTimes(1)
  })

  it('ignores an idle close from a socket the reconnect loop already replaced', async () => {
    const m = await freshModule()
    const drop = vi.fn<() => void>()
    m.installDropSocketHandler(drop)
    const obsolete = socket()
    m.noteSocketOpened(obsolete)
    m.noteSocketOpened(socket())

    m.noteSocketClosed(obsolete, SESSION_IDLE_CLOSE_CODE)
    expect(m.isConnectionPaused()).toBe(false)
    expect(drop).not.toHaveBeenCalled()
  })

  it('stays parked after an idle close until a deliberate resume', async () => {
    const m = await freshModule()
    m.installDropSocketHandler(() => {})
    const live = socket()
    m.noteSocketOpened(live)
    m.noteSocketClosed(live, SESSION_IDLE_CLOSE_CODE)

    // The parked waiter is the next dial; nothing but a resume releases it.
    let dialed = false
    void m.waitWhilePaused().then(() => { dialed = true })
    m.noteSocketClosed(live, SESSION_IDLE_CLOSE_CODE)
    await Promise.resolve()
    expect(dialed).toBe(false)

    m.resumeConnection()
    await Promise.resolve()
    expect(dialed).toBe(true)
  })

  it('ignores resume when not paused and can pause again after resuming', async () => {
    const m = await freshModule()
    const drop = vi.fn<() => void>()
    m.installDropSocketHandler(drop)

    m.resumeConnection()
    expect(m.isConnectionPaused()).toBe(false)

    m.pauseConnection()
    m.resumeConnection()
    m.pauseConnection()
    expect(drop).toHaveBeenCalledTimes(2)
    expect(m.isConnectionPaused()).toBe(true)
  })
})

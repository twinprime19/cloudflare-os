import { describe, expect, it, vi } from 'vitest'

// The module keeps its flag and waiters in module scope, the way `main.tsx` keeps the connection
// itself, so each test imports a fresh copy rather than sharing state through the suite.
async function freshModule() {
  vi.resetModules()
  return await import('./connectionPause')
}

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

  it('treats a server idle close as ordinary until a policy says otherwise', async () => {
    const m = await freshModule()
    expect(m.shouldPauseOnIdleClose()).toBe(false)

    m.installIdleClosePolicy(() => true)
    expect(m.shouldPauseOnIdleClose()).toBe(true)

    // A busy workspace keeps its policy installed but answers false, so the loop redials.
    m.installIdleClosePolicy(() => false)
    expect(m.shouldPauseOnIdleClose()).toBe(false)

    m.installIdleClosePolicy(null)
    expect(m.shouldPauseOnIdleClose()).toBe(false)
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

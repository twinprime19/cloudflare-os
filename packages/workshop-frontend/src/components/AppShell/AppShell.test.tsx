// @vitest-environment jsdom
/* eslint-disable react/react-in-jsx-scope */

import { act } from 'react'
import { type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'

vi.mock('@tanstack/react-router', () => ({ useRouterState: () => '/' }))

vi.mock('../../RpcContext', () => ({ useConnectionLost: () => false }))
vi.mock('../../TopBarNotice', () => ({ default: () => null }))
vi.mock('./CommandPalette', () => ({ default: () => null }))
vi.mock('./Sidebar', () => ({
  default: () => <aside data-testid="sidebar" />,
}))

import AppShell from './AppShell'
import { createI18nRoot } from '../../testI18nRoot'
import { isConnectionPaused, pauseConnection, resumeConnection } from '../../connectionPause'

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const PAUSED_LABEL = 'Paused — click to reconnect'

describe('AppShell', () => {
  let root: Root | undefined
  let container: HTMLDivElement | undefined

  function mount() {
    container = document.createElement('div')
    root = createI18nRoot(container)
    act(() => root!.render(<AppShell><div /></AppShell>))
  }

  // Every chip carries a label; the only other control in the shell's top bar is the icon-only
  // mobile hamburger.
  function chips() {
    return [...container!.querySelectorAll<HTMLElement>('button, [role="status"]')]
      .filter(el => el.textContent !== '')
  }

  afterEach(() => {
    act(() => root?.unmount())
    container?.remove()
    resumeConnection()
  })

  it('gives the percentage-height desktop sidebar a definite-height container', () => {
    mount()

    const sidebarContainer = container!.querySelector('[data-testid="sidebar"]')?.parentElement
    expect(sidebarContainer?.classList.contains('h-full')).toBe(true)
  })

  // The shell's routes have no idle timer, but the server parks them: a page holding an Overseer
  // stub keeps a session whose lease expires. Without this chip the parked reconnect loop has no
  // resume control outside the workspace editor.
  it('offers the paused chip while the connection is parked, and resumes on click', () => {
    pauseConnection()
    mount()

    const [chip, ...rest] = chips()
    expect(chip?.textContent).toBe(PAUSED_LABEL)
    expect(rest).toHaveLength(0)
    expect(chip.tagName).toBe('BUTTON')

    act(() => { chip.click() })
    expect(isConnectionPaused()).toBe(false)
    expect(chips()).toHaveLength(0)
  })

  it('shows no chip when the connection is neither paused nor lost', () => {
    mount()

    expect(chips()).toHaveLength(0)
  })
})

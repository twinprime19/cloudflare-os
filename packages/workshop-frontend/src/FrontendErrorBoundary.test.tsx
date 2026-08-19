// @vitest-environment jsdom
/* eslint-disable react/react-in-jsx-scope */

import { act } from 'react'
import { type Root } from 'react-dom/client'
import { createI18nRoot } from './testI18nRoot'
import { afterEach, describe, expect, it, vi } from 'vitest'
import FrontendErrorBoundary from './FrontendErrorBoundary'
import { reportIssue } from './errorReporting'

vi.mock('./errorReporting', () => ({ reportIssue: vi.fn<typeof reportIssue>() }))

function Broken(): never { throw new Error('render failed') }

describe('FrontendErrorBoundary', () => {
  let root: Root | undefined
  afterEach(() => { act(() => root?.unmount()); vi.restoreAllMocks() })

  it('reports a React crash and offers a reload action', async () => {
    vi.spyOn(console, 'error').mockImplementation(() => {})
    const container = document.createElement('div')
    document.body.append(container)
    root = createI18nRoot(container)
    await act(async () => root!.render(
      <FrontendErrorBoundary>
        <Broken />
      </FrontendErrorBoundary>,
    ))
    expect(container.textContent).toContain('Something went wrong')
    expect(container.querySelector('button')?.textContent).toContain('Reload')
    expect(reportIssue).toHaveBeenCalledWith('workshop.react-render', expect.any(Error),
      expect.objectContaining({ captureMechanism: 'react', handled: false }))
    container.remove()
  })
})

// @vitest-environment jsdom
/* eslint-disable react/react-in-jsx-scope */

import { act } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { RpcStub } from 'capnweb'
import type { AuthenticatedApi } from '@gadgets/workshop-shared/api'

const testState = vi.hoisted(() => ({
  userId: 'user-a',
  api: null as RpcStub<AuthenticatedApi> | null,
}))

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

vi.mock('@cloudflare/kumo', () => ({
  useKumoToastManager: () => ({ add: vi.fn<(toast: unknown) => void>() }),
}))

vi.mock('./AuthContext', () => ({
  useAuthenticatedApi: () => ({
    authenticatedApi: testState.api,
    currentUser: { id: testState.userId, name: 'User A' },
  }),
}))

vi.mock('./ThemeContext', () => ({ useTheme: () => ({ resolvedThemeMode: 'light' }) }))
vi.mock('./useDocumentTitle', () => ({ useDocumentTitle: () => {} }))
vi.mock('./AddModelModal', () => ({ default: () => null }))

// canvas/createImageBitmap are unavailable under jsdom; the rest of avatarUtils (the blob URL
// helper the wizard and the avatar cache both use) stays real.
vi.mock('./avatarUtils', async (importOriginal) => ({
  ...(await importOriginal<typeof import('./avatarUtils')>()),
  compressAvatar: async () => new Uint8Array([1, 2, 3]),
}))

import OnboardingWizard from './OnboardingWizard'

const originalCreateObjectURL = URL.createObjectURL
const originalRevokeObjectURL = URL.revokeObjectURL

const createdUrls: string[] = []
const revokedUrls: string[] = []

// Settles the effect chain: the avatar fetch and the compression step both resolve on microtasks.
async function flush() {
  await act(async () => {
    await Promise.resolve()
    await Promise.resolve()
    await Promise.resolve()
  })
}

function avatarImage(rendered: HTMLDivElement) {
  return rendered.querySelector<HTMLImageElement>('img[alt="Avatar preview"]')
}

function fakeApi(avatarBytes: Uint8Array | null): RpcStub<AuthenticatedApi> {
  return {
    listModels: async () => [],
    getAiConfig: async () => null,
    listGatekeeperVendors: async () => [],
    subscribeConnectedAccounts: () =>
      Object.assign(Promise.resolve({ [Symbol.dispose]() {} }), { [Symbol.dispose]() {} }),
    getAvatar: async () => avatarBytes,
  } as unknown as RpcStub<AuthenticatedApi>
}

describe('OnboardingWizard profile step', () => {
  let container: HTMLDivElement | undefined
  let root: Root | undefined

  beforeEach(() => {
    createdUrls.length = 0
    revokedUrls.length = 0
    URL.createObjectURL = vi.fn<(object: Blob | MediaSource) => string>(() => {
      const url = `blob:test/${createdUrls.length + 1}`
      createdUrls.push(url)
      return url
    })
    URL.revokeObjectURL = vi.fn<(url: string) => void>((url) => {
      revokedUrls.push(url)
    })
  })

  afterEach(async () => {
    await act(async () => root?.unmount())
    container?.remove()
    root = undefined
    container = undefined
    URL.createObjectURL = originalCreateObjectURL
    URL.revokeObjectURL = originalRevokeObjectURL
    vi.clearAllMocks()
  })

  async function render(userId: string, avatarBytes: Uint8Array | null) {
    testState.userId = userId
    testState.api = fakeApi(avatarBytes)
    container = document.createElement('div')
    document.body.append(container)
    root = createRoot(container)
    await act(async () => root!.render(<OnboardingWizard onComplete={() => {}} />))
    await flush()
    return container
  }

  async function pickFile(rendered: HTMLDivElement) {
    const input = rendered.querySelector<HTMLInputElement>('input[type="file"]')
    expect(input).not.toBeNull()
    Object.defineProperty(input, 'files', {
      value: [new File(['image-bytes'], 'me.png', { type: 'image/png' })],
      configurable: true,
    })
    await act(async () => {
      input!.dispatchEvent(new Event('change', { bubbles: true }))
    })
    await flush()
  }

  it('shows the avatar already stored for the user before any file is picked', async () => {
    const rendered = await render('seeded-user', new Uint8Array([9, 9, 9]))

    expect(createdUrls).toHaveLength(1)
    expect(avatarImage(rendered)?.getAttribute('src')).toBe(createdUrls[0])
    expect(rendered.textContent).toContain('Change')
  })

  it('keeps the empty camera placeholder when the user has no stored avatar', async () => {
    const rendered = await render('no-avatar-user', null)

    expect(avatarImage(rendered)).toBeNull()
    expect(rendered.textContent).toContain('Add photo')
  })

  it('leaves the shared avatar cache URL alive on unmount', async () => {
    await render('cache-owned-user', new Uint8Array([9, 9, 9]))
    const cacheOwnedUrl = createdUrls[0]

    await act(async () => root!.unmount())
    root = undefined

    expect(revokedUrls).not.toContain(cacheOwnedUrl)
    expect(revokedUrls).toEqual([])
  })

  it('renders a picked file over the stored avatar and revokes only that local URL', async () => {
    const rendered = await render('picked-file-user', new Uint8Array([9, 9, 9]))
    const cacheOwnedUrl = createdUrls[0]

    await pickFile(rendered)
    const localPreviewUrl = createdUrls[1]

    expect(localPreviewUrl).toBeDefined()
    expect(localPreviewUrl).not.toBe(cacheOwnedUrl)
    expect(avatarImage(rendered)?.getAttribute('src')).toBe(localPreviewUrl)

    await act(async () => root!.unmount())
    root = undefined

    expect(revokedUrls).toEqual([localPreviewUrl])
  })
})

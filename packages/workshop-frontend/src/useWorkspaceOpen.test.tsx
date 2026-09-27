// @vitest-environment jsdom
/* eslint-disable react/react-in-jsx-scope */

import { act, type ReactNode } from 'react'
import { type Root } from 'react-dom/client'
import { createI18nRoot } from './testI18nRoot'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { RpcStub } from 'capnweb'
import {
  createOpenGadgetError,
  OPEN_GADGET_ERROR_CODES,
  type AuthenticatedApi,
  type GadgetMetadata,
  type ObserverBindingNeed,
  type ObserverConfigCallback,
  type Overseer,
} from '@gadgets/workshop-shared/api'
import WorkspaceOpenErrorPage from './components/WorkspaceOpenErrorPage'
import {
  installDropSocketHandler,
  pauseConnection,
  resumeConnection,
} from './connectionPause'
import { reportIssue } from './errorReporting'
import { useWorkspaceOpen } from './useWorkspaceOpen'

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

vi.mock('./errorReporting', () => ({ reportIssue: vi.fn<typeof reportIssue>() }))

vi.mock('./components/WorkshopControls', () => ({
  WorkshopButton: ({ children, onClick }: { children: ReactNode; onClick?: () => void }) => (
    <button type="button" onClick={onClick}>{children}</button>
  ),
}))

// What capnweb rejects in-flight calls with when the pause disposes the main stub; the installed
// build's wording is pinned by rpcErrors.test.ts.
const DISPOSED_SESSION_MESSAGE = 'RPC session was shut down by disposing the main stub'

function deferred<T>() {
  let resolve!: (value: T) => void
  let reject!: (reason: unknown) => void
  const promise = new Promise<T>((next, fail) => { resolve = next; reject = fail })
  return { promise, resolve, reject }
}

function disposableStub<T extends object>(value: T, dispose = vi.fn<() => void>()) {
  return Object.assign(value, { [Symbol.dispose]: dispose }) as T & Disposable
}

function api(overseer: RpcStub<Overseer>): RpcStub<AuthenticatedApi> {
  return { openGadget: () => overseer } as unknown as RpcStub<AuthenticatedApi>
}

const METADATA = {
  id: 'workspace-1',
  title: 'Quarterly planning',
  provisional: false,
} as GadgetMetadata

function WorkspaceProbe({ authenticatedApi }: { authenticatedApi: RpcStub<AuthenticatedApi> }) {
  const state = useWorkspaceOpen({
    id: 'workspace-1',
    authenticatedApi,
    onInvalidShareKey: () => {},
    onMetadata: () => {},
    onShareKeyConsumed: () => {},
  })
  if (state.error?.kind === 'open') {
    return (
      <WorkspaceOpenErrorPage
        kind={state.error.failure}
        onGoToWorkspaces={() => {}}
        onRetry={state.retry}
      />
    )
  }
  if (state.error?.kind === 'message') return <p>{state.error.message}</p>
  return <p>{state.metadata?.title}</p>
}

describe('useWorkspaceOpen', () => {
  let root: Root | undefined
  let container: HTMLDivElement | undefined

  afterEach(() => {
    act(() => root?.unmount())
    container?.remove()
    document.title = ''
    act(() => resumeConnection())
    installDropSocketHandler(() => {})
    vi.mocked(reportIssue).mockClear()
    vi.restoreAllMocks()
  })

  it('disposes a metadata subscription that resolves after its load attempt is cleaned up', async () => {
    const pendingSubscription = deferred<RpcStub<{}>>()
    const overseerDispose = vi.fn<() => void>()
    const overseer = disposableStub({
      subscribeToMetadata: vi.fn<() => Promise<RpcStub<{}>>>(() => pendingSubscription.promise),
    }, overseerDispose) as unknown as RpcStub<Overseer>
    const subscriptionDispose = vi.fn<() => void>()
    const subscription = disposableStub({}, subscriptionDispose) as RpcStub<{}>
    const authenticatedApi = api(overseer)

    function Probe() {
      useWorkspaceOpen({
        id: 'workspace-1',
        authenticatedApi,
        onInvalidShareKey: () => {},
        onMetadata: () => {},
        onShareKeyConsumed: () => {},
      })
      return null
    }

    container = document.createElement('div')
    document.body.append(container)
    root = createI18nRoot(container)
    await act(async () => root!.render(<Probe />))

    act(() => root!.unmount())
    root = undefined
    await act(async () => { pendingSubscription.resolve(subscription); await Promise.resolve() })

    expect(overseerDispose).toHaveBeenCalledOnce()
    expect(subscriptionDispose).toHaveBeenCalledOnce()
  })

  it('clears loaded metadata and title and disposes the failed stub after access is denied', async () => {
    vi.spyOn(console, 'error').mockImplementation(() => {})
    document.title = 'outside'
    const firstSubscriptionDispose = vi.fn<() => void>()
    const firstOverseer = disposableStub({
      subscribeToMetadata: vi.fn<
        (callback: (metadata: GadgetMetadata) => void) => Promise<RpcStub<{}>>
      >(async callback => {
          callback(METADATA)
          return disposableStub({}, firstSubscriptionDispose) as RpcStub<{}>
        }),
    }) as unknown as RpcStub<Overseer>
    const deniedOverseerDispose = vi.fn<() => void>()
    const deniedOverseer = disposableStub({
      subscribeToMetadata: vi.fn<() => Promise<RpcStub<{}>>>(async () => {
        throw createOpenGadgetError(OPEN_GADGET_ERROR_CODES.workspaceAccessDenied)
      }),
    }, deniedOverseerDispose) as unknown as RpcStub<Overseer>

    container = document.createElement('div')
    document.body.append(container)
    root = createI18nRoot(container)
    await act(async () => root!.render(<WorkspaceProbe authenticatedApi={api(firstOverseer)} />))
    expect(container.textContent).toContain('Quarterly planning')
    expect(document.title).toBe('Quarterly planning - Cloudflare OS')

    await act(async () => root!.render(<WorkspaceProbe authenticatedApi={api(deniedOverseer)} />))
    expect(container.textContent).toContain("You don't have access to this workspace")
    expect(container.textContent).not.toContain('Quarterly planning')
    expect(document.title).toBe('Cloudflare OS')
    expect(firstSubscriptionDispose).toHaveBeenCalledOnce()
    expect(deniedOverseerDispose).toHaveBeenCalledOnce()
  })

  it('cancels an unanswered observer dialog on pause and leaves the dropped open unreported', async () => {
    // The open is still in flight while the dialog waits, so the pause drops the socket out from
    // under it: the failure belongs to the pause, not to the workspace.
    const pendingSubscription = deferred<RpcStub<{}>>()
    const overseer = disposableStub({
      subscribeToMetadata: vi.fn<() => Promise<RpcStub<{}>>>(() => pendingSubscription.promise),
    }) as unknown as RpcStub<Overseer>
    let configureObservers: RpcStub<ObserverConfigCallback> | undefined
    const authenticatedApi = {
      openGadget: (
        _id: string,
        _shareKey: string | undefined,
        callback: RpcStub<ObserverConfigCallback>,
      ) => {
        configureObservers = callback
        return overseer
      },
    } as unknown as RpcStub<AuthenticatedApi>

    let dialogOpen = false
    let loadError: ReturnType<typeof useWorkspaceOpen>['error'] = null
    function Probe() {
      const { observerConfig, error } = useWorkspaceOpen({
        id: 'workspace-1',
        authenticatedApi,
        onInvalidShareKey: () => {},
        onMetadata: () => {},
        onShareKeyConsumed: () => {},
      })
      dialogOpen = observerConfig !== null
      loadError = error
      return null
    }

    installDropSocketHandler(() => pendingSubscription.reject(new Error(DISPOSED_SESSION_MESSAGE)))
    container = document.createElement('div')
    document.body.append(container)
    root = createI18nRoot(container)
    await act(async () => root!.render(<Probe />))

    const need: ObserverBindingNeed = {
      gatekeeperId: 1,
      vendorId: 'google',
      resourceTitle: 'Quarterly planning',
    }
    let outcome!: Promise<string>
    await act(async () => {
      outcome = configureObservers!.configure([need]).then(
        () => 'answered',
        (error: unknown) => (error instanceof Error ? error.message : String(error)),
      )
      await Promise.resolve()
    })
    expect(dialogOpen).toBe(true)

    // The pause drops the socket the server is waiting on, so the dialog must not outlive it.
    await act(async () => { pauseConnection() })
    await expect(outcome).resolves.toContain('OBSERVER_CONFIG_CANCELLED')
    expect(dialogOpen).toBe(false)
    // The parked tab waits on the Paused chip and re-opens on resume, so the broken-session
    // rejection is neither an error page nor an issue worth reporting.
    expect(loadError).toBeNull()
    expect(reportIssue).not.toHaveBeenCalled()
  })

  it('shows the share-links-disabled page when a new redeemer is refused', async () => {
    vi.spyOn(console, 'error').mockImplementation(() => {})
    const overseer = disposableStub({
      subscribeToMetadata: vi.fn<() => Promise<RpcStub<{}>>>(async () => {
        throw createOpenGadgetError(OPEN_GADGET_ERROR_CODES.shareLinksDisabled)
      }),
    }) as unknown as RpcStub<Overseer>

    container = document.createElement('div')
    document.body.append(container)
    root = createI18nRoot(container)
    await act(async () => root!.render(<WorkspaceProbe authenticatedApi={api(overseer)} />))

    expect(container.textContent).toContain('Share links are turned off for this workspace')
    expect(container.textContent)
      .toContain('Ask the workspace owner to add you directly, then try again.')
  })
})

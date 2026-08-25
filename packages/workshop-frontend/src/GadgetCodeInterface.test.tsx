// @vitest-environment jsdom
/* eslint-disable react/react-in-jsx-scope */

import { act, type ComponentProps, type ReactElement, type ReactNode } from 'react'
import { type Root } from 'react-dom/client'
import { createI18nRoot } from './testI18nRoot'
import { afterAll, afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import * as Y from 'yjs'
import type { RpcStub } from 'capnweb'
import type { CodeSubscriber, Overseer } from '@gadgets/workshop-shared/api'

const testGlobal = globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
const previousActEnvironment = testGlobal.IS_REACT_ACT_ENVIRONMENT
testGlobal.IS_REACT_ACT_ENVIRONMENT = true
afterAll(() => {
  if (previousActEnvironment === undefined) delete testGlobal.IS_REACT_ACT_ENVIRONMENT
  else testGlobal.IS_REACT_ACT_ENVIRONMENT = previousActEnvironment
})

const mocks = vi.hoisted(() => ({
  downloadSourceZip: vi.fn<(files: Array<[string, string]>, gadgetTitle: string) => void>(),
  saveTextToFile: vi.fn<(filename: string, content: string) => void>(),
  toast: vi.fn<(toast: unknown) => void>(),
}))

vi.mock('@cloudflare/kumo', () => {
  const Dialog = Object.assign(
    ({ children }: { children: ReactNode }) => <div>{children}</div>,
    {
      Root: ({ children, open }: { children: ReactNode; open?: boolean }) => (
        open ? <div>{children}</div> : null
      ),
      Title: ({ children }: { children: ReactNode }) => <div>{children}</div>,
      Description: ({ children }: { children: ReactNode }) => <div>{children}</div>,
      Close: ({ render }: { render: (props: object) => ReactElement }) => render({}),
    },
  )
  const DropdownMenu = Object.assign(
    ({ children }: { children: ReactNode }) => <div>{children}</div>,
    {
      Trigger: ({ render }: { render: ReactElement }) => render,
      Content: ({ children }: { children: ReactNode }) => <div>{children}</div>,
      Item: ({ children, onClick }: { children: ReactNode; onClick?: () => void }) => (
        <button type="button" onClick={onClick}>{children}</button>
      ),
    },
  )
  return { Dialog, DropdownMenu, useKumoToastManager: () => ({ add: mocks.toast }) }
})

vi.mock('@phosphor-icons/react', () => {
  const icon = (name: string) => () => <span>{name}</span>
  return {
    DotsThree: icon('actions'),
    DownloadSimple: icon('download'),
    FileZip: icon('archive'),
    Pencil: icon('rename'),
    Plus: icon('new'),
    Trash: icon('delete'),
    X: icon('close'),
  }
})

vi.mock('./components/WorkshopControls', () => ({
  WorkshopButton: ({ children, ...props }: ComponentProps<'button'>) => (
    <button type="button" {...props}>{children}</button>
  ),
  WorkshopIconButton: ({ children, ...props }: ComponentProps<'button'>) => (
    <button type="button" {...props}>{children}</button>
  ),
  WorkshopInput: (props: ComponentProps<'input'>) => <input {...props} />,
}))

vi.mock('./components/DeleteConfirmationDialog', () => ({ default: () => null }))
vi.mock('./CodeEditor', () => ({ default: () => <div>editor</div> }))
vi.mock('./CodeDiffEditor', () => ({ default: () => <div>diff editor</div> }))

vi.mock('./fileTransfers', () => ({
  downloadSourceZip: mocks.downloadSourceZip,
  saveTextToFile: mocks.saveTextToFile,
}))

import GadgetCodeInterface from './GadgetCodeInterface'

const FILES_ROOT = 'files'

let container: HTMLDivElement
let root: Root

beforeEach(() => {
  mocks.downloadSourceZip.mockReset()
  mocks.saveTextToFile.mockReset()
  mocks.toast.mockReset()
  container = document.createElement('div')
  document.body.appendChild(container)
  root = createI18nRoot(container)
})

afterEach(async () => {
  await act(async () => root.unmount())
  container.remove()
})

/**
 * The server side of the code subscription: hands back the subscriber the component registers, so
 * a test can push the committed files in exactly as the backend would.
 */
function fakeOverseer(): { stub: RpcStub<Overseer>; subscriber: () => CodeSubscriber } {
  let registered: CodeSubscriber | null = null
  const stub = {
    subscribeToCode: vi.fn<(subscriber: CodeSubscriber) => Promise<Disposable>>(
      async (subscriber) => {
        registered = subscriber
        return { [Symbol.dispose]() {} }
      },
    ),
    updateCode: vi.fn<(update: Uint8Array, chatId?: number) => Promise<void>>(async () => {}),
  }
  return {
    stub: stub as unknown as RpcStub<Overseer>,
    subscriber: () => {
      if (!registered) throw new Error('the component never subscribed to code updates')
      return registered
    },
  }
}

/** A Yjs update that puts `files` into the committed map, as the server's first update does. */
function committedFilesUpdate(files: Record<string, string>): Uint8Array {
  const doc = new Y.Doc()
  const map = doc.getMap<Y.Text>(FILES_ROOT)
  for (const [filename, content] of Object.entries(files)) {
    map.set(filename, new Y.Text(content))
  }
  return Y.encodeStateAsUpdateV2(doc)
}

/**
 * The `proposedChanges` update a chat's branch would carry: `edits` applied on top of the committed
 * state, where a `null` content deletes the file. Built from the very bytes the server sends as the
 * committed update, so the two share struct ids and a deletion in here actually deletes over there.
 */
function proposedChangesUpdate(
  committed: Uint8Array,
  edits: Record<string, string | null>,
): Uint8Array {
  const doc = new Y.Doc()
  Y.applyUpdateV2(doc, committed)
  const map = doc.getMap<Y.Text>(FILES_ROOT)
  for (const [filename, content] of Object.entries(edits)) {
    if (content === null) map.delete(filename)
    else map.set(filename, new Y.Text(content))
  }
  return Y.encodeStateAsUpdateV2(doc)
}

/**
 * Mounts the editor over `files` and syncs them in as the server would.
 *
 * Pass `proposal` to mount in branch mode instead: it receives the committed update and returns the
 * branch's proposed changes, which the component then shows as the unaccepted diff.
 */
async function mountEditor(
  files: Record<string, string>,
  gadgetTitle = 'Gadget',
  proposal?: (committed: Uint8Array) => Uint8Array,
) {
  const overseer = fakeOverseer()
  const committed = Object.keys(files).length > 0 ? committedFilesUpdate(files) : null
  const proposedChanges = committed && proposal ? proposal(committed) : undefined

  await act(async () => {
    root.render(
      <GadgetCodeInterface
        overseer={overseer.stub}
        filesRoot={FILES_ROOT}
        gadgetTitle={gadgetTitle}
        isAgentActive={false}
        selectedChatId={proposedChanges ? 1 : null}
        proposedChanges={proposedChanges}
      />,
    )
  })

  await act(async () => {
    const subscriber = overseer.subscriber()
    if (committed) {
      subscriber.update({
        version: 1,
        timestamp: new Date(),
        update: committed,
      })
    }
    subscriber.ready()
  })

  return overseer
}

function downloadSourceButton(): HTMLButtonElement | null {
  return container.querySelector<HTMLButtonElement>('[aria-label="Download source"]')
}

describe('GadgetCodeInterface source download', () => {
  it('offers no source download while the gadget has no files', async () => {
    await mountEditor({})

    expect(container.textContent).toContain('No files yet')
    expect(downloadSourceButton()).toBeNull()
  })

  it('downloads every committed file under the gadget title', async () => {
    await mountEditor({
      'server.js': 'export default {}',
      'README.md': '# Gadget',
      'docs/notes.md': 'Ghi chú',
    }, 'Báo cáo tuần')

    await act(async () => { downloadSourceButton()?.click() })

    expect(mocks.downloadSourceZip).toHaveBeenCalledTimes(1)
    expect(mocks.downloadSourceZip).toHaveBeenCalledWith(
      [
        ['README.md', '# Gadget'],
        ['docs/notes.md', 'Ghi chú'],
        ['server.js', 'export default {}'],
      ],
      'Báo cáo tuần',
    )
  })

  it('downloads the proposed files on screen, without the one the proposal deletes', async () => {
    await mountEditor(
      {
        'README.md': '# Gadget',
        'server.js': 'export default {}',
        'legacy.js': 'superseded',
      },
      'Gadget',
      committed => proposedChangesUpdate(committed, {
        'server.js': 'export default { proposed: true }',
        'client.js': 'added on the branch',
        'legacy.js': null,
      }),
    )

    await act(async () => { downloadSourceButton()?.click() })

    expect(mocks.downloadSourceZip).toHaveBeenCalledTimes(1)
    expect(mocks.downloadSourceZip).toHaveBeenCalledWith(
      [
        ['README.md', '# Gadget'],
        ['client.js', 'added on the branch'],
        ['server.js', 'export default { proposed: true }'],
      ],
      'Gadget',
    )
  })

  it('offers no source download when the proposal deletes every file', async () => {
    await mountEditor(
      { 'server.js': 'export default {}' },
      'Gadget',
      committed => proposedChangesUpdate(committed, { 'server.js': null }),
    )

    // The name is still listed, struck through, so the user can see the deletion -- but there is
    // nothing left to archive, and a control that can only produce an empty file is worse than none.
    expect(container.textContent).toContain('server.js')
    expect(downloadSourceButton()).toBeNull()
  })

  it('reports a failed archive instead of throwing at the click', async () => {
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => {})
    mocks.downloadSourceZip.mockImplementation(() => { throw new Error('archive failed') })

    try {
      await mountEditor({ 'server.js': 'export default {}' })
      await act(async () => { downloadSourceButton()?.click() })

      expect(mocks.toast).toHaveBeenCalledWith(
        expect.objectContaining({ variant: 'error' }),
      )
    } finally {
      consoleError.mockRestore()
    }
  })
})

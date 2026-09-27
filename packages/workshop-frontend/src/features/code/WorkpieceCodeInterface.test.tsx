// @vitest-environment jsdom
/* eslint-disable react/react-in-jsx-scope */

import { act, type ComponentProps, type ReactElement, type ReactNode } from 'react'
import { type Root } from 'react-dom/client'
import { createI18nRoot } from '../../testI18nRoot'
import { afterAll, afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { RpcStub } from 'capnweb'
import type {
  CodeChangeSubmission, FileAtCommit, Overseer, TreeNode, WorkpieceId,
} from '@gadgets/workshop-shared/api'
import type { CodeChange, FileChange } from '@gadgets/workshop-shared/code-change'
import type { ChatCodeChanges } from '../../ChatInterface'

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
  const Collapsible = {
    Root: ({ children }: { children: ReactNode }) => <div>{children}</div>,
    Trigger: ({ children }: { children: ReactNode }) => <button type="button">{children}</button>,
    Panel: ({ children }: { children: ReactNode }) => <div>{children}</div>,
  }
  return { Collapsible, Dialog, DropdownMenu, useKumoToastManager: () => ({ add: mocks.toast }) }
})

vi.mock('@phosphor-icons/react', () => {
  const icon = (name: string) => () => <span>{name}</span>
  return {
    CaretDown: icon('collapse'),
    CaretRight: icon('expand'),
    Cube: icon('submodule'),
    DotsThree: icon('actions'),
    DownloadSimple: icon('download'),
    File: icon('file'),
    FileZip: icon('archive'),
    FolderSimple: icon('folder'),
    GitBranch: icon('branch'),
    LinkSimple: icon('symlink'),
    List: icon('files'),
    Pencil: icon('rename'),
    Plus: icon('new'),
    TerminalWindow: icon('executable'),
    Trash: icon('delete'),
    X: icon('close'),
  }
})

vi.mock('../../components/WorkshopControls', () => ({
  WorkshopButton: ({ children, ...props }: ComponentProps<'button'>) => (
    <button type="button" {...props}>{children}</button>
  ),
  WorkshopIconButton: ({ children, ...props }: ComponentProps<'button'>) => (
    <button type="button" {...props}>{children}</button>
  ),
  WorkshopInput: (props: ComponentProps<'input'>) => <input {...props} />,
}))

vi.mock('../../components/DeleteConfirmationDialog', () => ({ default: () => null }))
vi.mock('./CodeEditor', () => ({ default: () => <div>editor</div> }))
vi.mock('./CodeDiffEditor', () => ({ default: () => <div>diff editor</div> }))

vi.mock('../../fileTransfers', () => ({
  downloadSourceZip: mocks.downloadSourceZip,
  saveTextToFile: mocks.saveTextToFile,
}))

import WorkpieceCodeInterface from './WorkpieceCodeInterface'

const GADGET_ID = 1 as WorkpieceId
const CHAT_ID = 1

// Mints the fresh commit id each mount needs (see fakeOverseer).
let commitCounter = 1

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

/** The nested tree Overseer.listTree() answers for a flat path -> text map. */
function treeOf(files: Record<string, string>): TreeNode[] {
  const roots: TreeNode[] = []
  for (const path of Object.keys(files)) {
    const parts = path.split('/')
    let level = roots
    for (const dirName of parts.slice(0, -1)) {
      let dir = level.find(node => node.kind === 'dir' && node.name === dirName)
      if (dir === undefined) {
        dir = { name: dirName, kind: 'dir', children: [] }
        level.push(dir)
      }
      level = (dir as Extract<TreeNode, { kind: 'dir' }>).children
    }
    level.push({ name: parts[parts.length - 1]!, kind: 'file' })
  }
  return roots
}

/**
 * The server side of the code view: hands back the committed trees and file contents the
 * component asks for by commit id, lazily, the way commitFileStore reads them.
 *
 * Commits are immutable, so commitFileStore memoizes trees and contents per commit id for the
 * page's lifetime -- a page-wide singleton this test file shares across its cases, with no reset.
 * Every mount therefore mints a fresh commit id (see `mountEditor`) so a case really reads its own
 * files.
 */
function fakeOverseer(commits: Record<string, Record<string, string>>): RpcStub<Overseer> {
  const committed = (commitId: string) => {
    const files = commits[commitId]
    if (!files) throw new Error(`no such commit: ${commitId}`)
    return files
  }
  const stub = {
    listTree: vi.fn<(commitId: string) => Promise<TreeNode[]>>(
      async (commitId) => treeOf(committed(commitId)),
    ),
    readFilesAtCommit: vi.fn<
      (commitId: string, paths: string[]) => Promise<[string, FileAtCommit][]>
    >(async (commitId, paths) => {
      const files = committed(commitId)
      return paths.map(path => [
        path,
        Object.hasOwn(files, path) ? { kind: 'text', text: files[path]! } : { kind: 'absent' },
      ])
    }),
    submitCodeChange: vi.fn<
      (chatId: number, submission: CodeChangeSubmission) =>
        Promise<{ generation: number; revision: number }>
    >(async () => ({ generation: 0, revision: 0 })),
  }
  return stub as unknown as RpcStub<Overseer>
}

/**
 * The chat's uncommitted changes to the gadget, as its durable code state carries them: the
 * gadget pinned at the committed tree the test mounted, plus one recorded change per entry, where
 * a `null` content deletes the file.
 */
function chatCodeChanges(
  baseCommit: string, changes: Record<string, string | null>,
): ChatCodeChanges {
  const fileChanges: [string, FileChange][] = Object.entries(changes).map(
    ([filename, content]) => [filename, content === null ? { remove: true } : { set: content }],
  )
  const epochChange: CodeChange = { [GADGET_ID]: fileChanges }
  return {
    chatId: CHAT_ID,
    codeBase: {
      pins: [{ gadgetId: GADGET_ID, baseCommit, mergedCommit: baseCommit }],
      generation: 0,
      revision: 0,
    },
    epochChange,
    rowsThrough: 0,
  }
}

/** Lets the committed-tree fetch, the OT client's queue, and the coalesced re-renders land. */
async function settle() {
  await act(async () => { await new Promise(resolve => setTimeout(resolve, 50)) })
}

/**
 * Mounts the code view over `files` as the gadget's committed head.
 *
 * Pass `changes` to mount inside a chat instead: they become the chat's uncommitted changes on top
 * of those committed files, which the view then shows as the unaccepted diff.
 */
async function mountEditor(
  files: Record<string, string>,
  { title = 'Gadget', changes }: {
    title?: string
    changes?: Record<string, string | null>
  } = {},
) {
  const headCommitId = `head-${commitCounter++}`
  const overseer = fakeOverseer({ [headCommitId]: files })
  const chatChanges = changes && chatCodeChanges(headCommitId, changes)

  await act(async () => {
    root.render(
      <WorkpieceCodeInterface
        overseer={overseer}
        summary={{ id: GADGET_ID, type: 'gadget', title, commitId: headCommitId }}
        isAgentActive={false}
        selectedChatId={chatChanges ? CHAT_ID : null}
        chatChanges={chatChanges}
      />,
    )
  })
  await settle()

  return overseer
}

function sortedByPath(files: Array<[string, string]>): Array<[string, string]> {
  return files.toSorted(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
}

function downloadSourceButton(): HTMLButtonElement | null {
  return container.querySelector<HTMLButtonElement>('[aria-label="Download source"]')
}

describe('WorkpieceCodeInterface source download', () => {
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
    }, { title: 'Báo cáo tuần' })

    await act(async () => { downloadSourceButton()?.click() })

    // The archive follows the tree's walk (directories before files), so compare by path.
    expect(mocks.downloadSourceZip).toHaveBeenCalledTimes(1)
    const [files, title] = mocks.downloadSourceZip.mock.calls[0]!
    expect(sortedByPath(files)).toEqual([
      ['README.md', '# Gadget'],
      ['docs/notes.md', 'Ghi chú'],
      ['server.js', 'export default {}'],
    ])
    expect(title).toBe('Báo cáo tuần')
  })

  it('downloads the files on screen, without the one the chat deletes', async () => {
    await mountEditor(
      {
        'README.md': '# Gadget',
        'server.js': 'export default {}',
        'legacy.js': 'superseded',
      },
      {
        changes: {
          'server.js': 'export default { proposed: true }',
          'client.js': 'added on the branch',
          'legacy.js': null,
        },
      },
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

  it('offers no source download when the chat deletes every file', async () => {
    await mountEditor(
      { 'server.js': 'export default {}' },
      { changes: { 'server.js': null } },
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

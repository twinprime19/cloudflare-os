// @vitest-environment jsdom

import { afterEach, describe, expect, it, vi } from 'vitest'
import { strFromU8, unzipSync } from 'fflate'
import { downloadSourceZip, makeExportFilename, saveStreamToFile } from './fileTransfers'

afterEach(() => {
  delete (window as Window & { showSaveFilePicker?: unknown }).showSaveFilePicker
})

/**
 * Runs `download` with the browser's download plumbing stubbed, and returns what it saved.
 *
 * jsdom has neither object URLs nor a navigation to follow the anchor to, so both are replaced for
 * the call: the blob is taken where it is turned into a URL, and the file name off the anchor that
 * would have carried it to disk.
 */
async function captureDownload(download: () => void): Promise<{ filename: string; blob: Blob }> {
  const blobs: Blob[] = []
  const filenames: string[] = []
  const oldCreateObjectURL = URL.createObjectURL
  const oldRevokeObjectURL = URL.revokeObjectURL
  URL.createObjectURL = vi.fn<(blob: Blob) => string>(blob => {
    blobs.push(blob)
    return 'blob:test'
  })
  URL.revokeObjectURL = vi.fn<(url: string) => void>()
  const click = vi.spyOn(HTMLAnchorElement.prototype, 'click')
    .mockImplementation(function (this: HTMLAnchorElement) {
      filenames.push(this.download)
    })

  try {
    download()
  } finally {
    URL.createObjectURL = oldCreateObjectURL
    URL.revokeObjectURL = oldRevokeObjectURL
    click.mockRestore()
  }

  expect(blobs).toHaveLength(1)
  expect(filenames).toHaveLength(1)
  return { filename: filenames[0], blob: blobs[0] }
}

// `FileReader` rather than `blob.arrayBuffer()`: jsdom's Blob does not implement the latter.
function readBlob(blob: Blob): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.addEventListener('error', () => {
      reject(reader.error ?? new Error('could not read the downloaded blob'))
    })
    reader.addEventListener('load', () => {
      resolve(new Uint8Array(reader.result as ArrayBuffer))
    })
    reader.readAsArrayBuffer(blob)
  })
}

async function unzipDownload(blob: Blob): Promise<Record<string, string>> {
  const archive = unzipSync(await readBlob(blob))
  return Object.fromEntries(
    Object.entries(archive).map(([path, bytes]) => [path, strFromU8(bytes)]),
  )
}

describe('export file transfers', () => {
  it('builds a safe filename with the advertised extension', () => {
    expect(makeExportFilename('Quarterly report / 2026', '.csv'))
      .toBe('Quarterly-report-2026.csv')
  })

  it('opens the picker before starting the export stream', async () => {
    const order: string[] = []
    const picker = vi.fn<(_options: unknown) => Promise<{
      createWritable(): Promise<WritableStream<Uint8Array>>
    }>>(async () => {
      order.push('picker')
      return {
        async createWritable() {
          order.push('writable')
          return new WritableStream<Uint8Array>({
            write(bytes) {
              order.push(`write:${new TextDecoder().decode(bytes)}`)
            },
            close() {
              order.push('close')
            },
          })
        },
      }
    })
    Object.assign(window, { showSaveFilePicker: picker })

    await saveStreamToFile(
      async () => {
        order.push('source')
        return new ReadableStream<Uint8Array>({
          start(controller) {
            controller.enqueue(new TextEncoder().encode('value'))
            controller.close()
          },
        })
      },
      'report.csv',
      { description: 'CSV', contentType: 'text/csv', extension: '.csv' },
    )

    expect(order).toEqual(['picker', 'writable', 'source', 'write:value', 'close'])
    expect(picker).toHaveBeenCalledWith({
      suggestedName: 'report.csv',
      types: [{ description: 'CSV', accept: { 'text/csv': ['.csv'] } }],
    })
  })

  it('does not start a lazy export when the file picker is cancelled', async () => {
    Object.assign(window, {
      showSaveFilePicker: vi.fn<() => Promise<never>>()
        .mockRejectedValue(new DOMException('Cancelled', 'AbortError')),
    })
    const source = vi.fn<() => Promise<ReadableStream<Uint8Array>>>()

    await expect(saveStreamToFile(
      source,
      'report.pdf',
      { description: 'PDF', contentType: 'application/pdf', extension: '.pdf' },
    )).resolves.toBeUndefined()

    expect(source).not.toHaveBeenCalled()
  })

  it('propagates file picker failures other than cancellation', async () => {
    Object.assign(window, {
      showSaveFilePicker: vi.fn<() => Promise<never>>().mockRejectedValue(new Error('picker failed')),
    })
    const source = vi.fn<() => Promise<ReadableStream<Uint8Array>>>()

    await expect(saveStreamToFile(
      source,
      'report.pdf',
      { description: 'PDF', contentType: 'application/pdf', extension: '.pdf' },
    )).rejects.toThrow('picker failed')

    expect(source).not.toHaveBeenCalled()
  })

  it('preserves the advertised content type in the Blob fallback', async () => {
    const createObjectURL = vi.fn<(blob: Blob) => string>(() => 'blob:test')
    const oldCreateObjectURL = URL.createObjectURL
    const oldRevokeObjectURL = URL.revokeObjectURL
    const click = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {})
    URL.createObjectURL = createObjectURL
    URL.revokeObjectURL = vi.fn<(url: string) => void>()
    try {
      await saveStreamToFile(
        async () => new ReadableStream({
          start(controller) {
            controller.enqueue(new TextEncoder().encode('a,b'))
            controller.close()
          },
        }),
        'report.csv',
        { description: 'CSV', contentType: 'text/csv', extension: '.csv' },
      )

      expect(createObjectURL.mock.calls[0]?.[0].type).toBe('text/csv')
    } finally {
      URL.createObjectURL = oldCreateObjectURL
      URL.revokeObjectURL = oldRevokeObjectURL
      click.mockRestore()
    }
  })

})

describe('source zip download', () => {
  it('archives every file under its own name, subfolders included', async () => {
    const { filename, blob } = await captureDownload(() => downloadSourceZip([
      ['server.js', 'export default {}'],
      ['README.md', '# Gadget\n'],
      ['docs/notes.md', 'Ghi chú của nhóm'],
    ], 'Gadget'))

    expect(filename).toBe('Gadget.zip')
    expect(blob.type).toBe('application/zip')
    expect(await unzipDownload(blob)).toEqual({
      'server.js': 'export default {}',
      'README.md': '# Gadget\n',
      'docs/notes.md': 'Ghi chú của nhóm',
    })
  })

  it('drops names an archive cannot carry and keeps the rest', async () => {
    const { blob } = await captureDownload(() => downloadSourceZip([
      ['/server.js', 'rooted'],
      ['../escape.js', 'traversal'],
      ['docs/../escape.js', 'traversal'],
      ['', 'nameless'],
      ['/', 'separators only'],
      ['client.js', 'kept'],
    ], 'Gadget'))

    expect(await unzipDownload(blob)).toEqual({
      'server.js': 'rooted',
      'client.js': 'kept',
    })
  })

  it('drops a backslash traversal, which a Windows extractor would follow out of the folder', async () => {
    const { blob } = await captureDownload(() => downloadSourceZip([
      ['..\\..\\evil.bat', 'traversal'],
      ['docs\\..\\escape.js', 'traversal'],
      ['client.js', 'kept'],
    ], 'Gadget'))

    expect(await unzipDownload(blob)).toEqual({ 'client.js': 'kept' })
  })

  it('archives a backslash path as the folder its name means', async () => {
    const { blob } = await captureDownload(
      () => downloadSourceZip([['docs\\a.md', 'Ghi chú']], 'Gadget'),
    )

    expect(await unzipDownload(blob)).toEqual({ 'docs/a.md': 'Ghi chú' })
  })

  it('drops a drive-lettered absolute path, however it is dressed up', async () => {
    const { blob } = await captureDownload(() => downloadSourceZip([
      ['C:\\abs.txt', 'absolute'],
      // All three reduce to the same `C:/evil.bat`, which is why the drive check has to run on the
      // reduced segments rather than on the name as it arrived.
      ['./C:/evil.bat', 'absolute'],
      ['.\\C:\\evil.bat', 'absolute'],
      ['/C:/evil.bat', 'absolute'],
      ['client.js', 'kept'],
    ], 'Gadget'))

    expect(await unzipDownload(blob)).toEqual({ 'client.js': 'kept' })
  })

  it('drops `__proto__` rather than letting it corrupt the whole archive', async () => {
    const { blob } = await captureDownload(() => downloadSourceZip([
      ['README.md', 'keep'],
      ['__proto__', 'pwn'],
    ], 'Gadget'))

    // An exact match, not a `README.md` lookup: as an entry key `__proto__` sets the entry object's
    // prototype, and the archive that comes out carries thousands of junk entries, not one.
    expect(await unzipDownload(blob)).toEqual({ 'README.md': 'keep' })
  })

  it('collapses the separator noise an extractor would collapse anyway, last name wins', async () => {
    const { blob } = await captureDownload(() => downloadSourceZip([
      ['a//b.md', 'first'],
      ['a/./b.md', 'second'],
    ], 'Gadget'))

    // Both names extract to `a/b.md`, so keeping them apart in the archive would only decide which
    // one survives at extraction time. They collapse here instead, and the later file wins.
    expect(await unzipDownload(blob)).toEqual({ 'a/b.md': 'second' })
  })

  it('keeps a name that ends in a separator as a file, with its content', async () => {
    const { blob } = await captureDownload(
      () => downloadSourceZip([['trailing./', 'content']], 'Gadget'),
    )

    expect(await unzipDownload(blob)).toEqual({ 'trailing.': 'content' })
  })

  it('saves the archive as one blob without opening the file picker', async () => {
    const picker = vi.fn<() => Promise<never>>()
    Object.assign(window, { showSaveFilePicker: picker })

    const { blob } = await captureDownload(() => downloadSourceZip([['server.js', 'ok']], 'Gadget'))

    expect(picker).not.toHaveBeenCalled()
    expect(blob.type).toBe('application/zip')
  })

  it('names the archive after the gadget, transliterated', async () => {
    const { filename } = await captureDownload(
      () => downloadSourceZip([['server.js', 'ok']], 'Báo cáo tuần'),
    )

    expect(filename).toBe('Bao-cao-tuan.zip')
  })
})

describe('archive filenames', () => {
  it('keeps Vietnamese titles readable', () => {
    expect(makeExportFilename('Đơn hàng', '.zip')).toBe('Don-hang.zip')
    expect(makeExportFilename('Chi phí & doanh thu', '.zip')).toBe('Chi-phi-doanh-thu.zip')
  })

  it('falls back when a title has no ASCII left to keep', () => {
    expect(makeExportFilename('報告', '.zip')).toBe('gadget.zip')
  })
})

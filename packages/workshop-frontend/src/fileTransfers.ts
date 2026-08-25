import { strToU8, zipSync } from 'fflate'

export const BLUEPRINT_ARCHIVE_EXTENSION = '.gadget'

const SOURCE_ARCHIVE_EXTENSION = '.zip'

function makeFilename(title: string, fallback: string): string {
  // Decomposing first turns an accented Latin letter into its base letter plus a combining mark, so
  // the strip below keeps the letter instead of collapsing the whole word into separators: a
  // Vietnamese title reaches the filesystem readable ("Báo cáo" -> "Bao-cao") rather than as
  // "-" runs. `đ`/`Đ` carry the stroke in the letter itself and have no decomposition, so they are
  // mapped by hand. A title with nothing Latin left (CJK, for one) still falls back.
  return title
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .trim()
    .replace(/[^A-Za-z0-9._-]+/g, '-')
    .replace(/^-+|-+$/g, '') || fallback
}

export function makeBlueprintFilename(title: string, version: number): string {
  return `${makeFilename(title, 'blueprint')}-v${version}${BLUEPRINT_ARCHIVE_EXTENSION}`
}

export function makeExportFilename(title: string, extension: string): string {
  return `${makeFilename(title, 'gadget')}${extension}`
}

type SaveFileHandle = {
  createWritable(): Promise<WritableStream<Uint8Array>>
}

type SaveFilePicker = (options: {
  suggestedName?: string
  types?: Array<{
    description?: string
    accept: Record<string, string[]>
  }>
}) => Promise<SaveFileHandle>

function triggerBlobDownload(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob)

  try {
    const link = document.createElement('a')
    link.href = url
    link.download = filename
    link.style.display = 'none'
    document.body.appendChild(link)
    link.click()
    link.remove()
  } finally {
    window.setTimeout(() => URL.revokeObjectURL(url), 100)
  }
}

export async function saveStreamToFile(
  createStream: () => Promise<ReadableStream<Uint8Array>>,
  filename: string,
  fileType: {
    description: string
    contentType: string
    extension: string
  },
): Promise<void> {
  const showSaveFilePicker = (window as Window & {
    showSaveFilePicker?: SaveFilePicker
  }).showSaveFilePicker

  if (showSaveFilePicker) {
    // Open the file picker immediately after user interaction and before fetching file stream
    // to avoid browser security errors raised when delay is too long.
    let handle: SaveFileHandle
    try {
      handle = await showSaveFilePicker({
        suggestedName: filename,
        types: [{
          description: fileType.description,
          accept: {
            [fileType.contentType]: [fileType.extension],
          },
        }],
      })
    } catch (error) {
      // AbortError means the user exited the file picker without selecting a destination.
      if (!(error instanceof DOMException) || error.name !== 'AbortError') throw error
      return
    }

    const writable = await handle.createWritable()
    let stream: ReadableStream<Uint8Array>
    try {
      stream = await createStream()
    } catch (error) {
      await writable.abort(error).catch(() => {})
      throw error
    }
    await stream.pipeTo(writable)
    return
  }

  const stream = await createStream()
  triggerBlobDownload(await new Response(stream, {
    headers: { 'Content-Type': fileType.contentType },
  }).blob(), filename)
}

export function saveTextToFile(filename: string, content: string): void {
  triggerBlobDownload(new Blob([content], { type: 'text/plain;charset=utf-8' }), filename)
}

/**
 * The archive path a gadget file name gets, or `null` when the file cannot be archived at all.
 *
 * Names are kept verbatim otherwise -- they are the gadget's own file names, so a path-like name
 * (`docs/notes.md`) has to extract as a folder. Only the two shapes no archive may carry are
 * rejected: a name that is empty once leading separators are gone, and one that walks out of the
 * archive root, which some extractors follow.
 */
function sourceEntryPath(filename: string): string | null {
  const path = filename.replace(/^\/+/, '')
  if (!path) return null
  if (path.split('/').includes('..')) return null
  return path
}

/**
 * Downloads a gadget's source files as a single zip named after the gadget.
 *
 * Assembled in the browser and saved as one blob, never through the save-file picker: the archive
 * is already fully in memory by the time it is offered, so there is no stream to keep open, and
 * the picker's user-gesture window is the one thing that could still fail here.
 */
export function downloadSourceZip(files: Array<[string, string]>, gadgetTitle: string): void {
  const entries: Record<string, Uint8Array> = {}
  for (const [filename, content] of files) {
    const path = sourceEntryPath(filename)
    if (path === null) continue
    entries[path] = strToU8(content)
  }

  const archive = zipSync(entries)
  triggerBlobDownload(
    new Blob([archive], { type: 'application/zip' }),
    makeExportFilename(gadgetTitle, SOURCE_ARCHIVE_EXTENSION),
  )
}

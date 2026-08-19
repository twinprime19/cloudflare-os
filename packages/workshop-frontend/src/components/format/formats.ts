// How an output format is drawn: which glyph, which nouns, which wireframe. The server sends a
// blueprint's `BlueprintOutput` declaration; this module is the only place that maps it to pixels.
//
// Icons are a closed set (`OUTPUT_ICONS` in the shared API) so a free-form name can't make a
// deployment look wrong. Only the key crosses the wire; the glyphs live here.

import {
  FileText,
  GridNine,
  Presentation,
  AppWindow,
  FlowArrow,
  Kanban,
  ChartBar,
  Table,
  Notebook,
  ListChecks,
  type Icon as PhosphorIcon,
} from '@phosphor-icons/react'
import { msg } from '@lingui/core/macro'
import { i18n } from '@gadgets/i18n'
import type { BlueprintOutput, OutputIcon } from '@gadgets/workshop-shared/api'

/** The glyph for each key in the shared `OUTPUT_ICONS` vocabulary. */
export const FORMAT_ICONS = {
  fileText: FileText,
  gridNine: GridNine,
  presentation: Presentation,
  appWindow: AppWindow,
  flowArrow: FlowArrow,
  kanban: Kanban,
  chartBar: ChartBar,
  table: Table,
  notebook: Notebook,
  listChecks: ListChecks,
} satisfies Record<OutputIcon, PhosphorIcon>

/**
 * Which wireframe illustrates a format. Derived from the icon rather than picked separately, so
 * the two can't disagree; icons depicting the same artefact (a page vs. a notebook) share one.
 */
export type FormatWireframe = 'page' | 'grid' | 'slide' | 'window' | 'list' | 'board' | 'chart'

const WIREFRAME_FOR_ICON: Record<OutputIcon, FormatWireframe> = {
  fileText: 'page',
  notebook: 'page',
  gridNine: 'grid',
  table: 'grid',
  presentation: 'slide',
  chartBar: 'chart',
  appWindow: 'window',
  flowArrow: 'board',
  kanban: 'board',
  listChecks: 'list',
}

// Message descriptors, not resolved strings: `GENERIC_OUTPUT` below is a module-scope constant
// evaluated once at import time, before the active locale is guaranteed to be set. `msg` only
// captures the source text and message id at compile time, so building the descriptor here has no
// dependency on locale. `GENERIC_OUTPUT.noun`/`.plural` resolve it through a getter instead of
// storing it as a plain string, so every read re-resolves against whichever locale is active at
// that moment, the same way `BlueprintOutput.noun`/`.plural` do for every other (server-declared)
// output -- while still satisfying `BlueprintOutput`'s plain `string` fields for every consumer.
const GENERIC_OUTPUT_NOUN = msg`App`
const GENERIC_OUTPUT_PLURAL = msg`Apps`

/**
 * How a gadget with no declared format is shown. Also the fallback for an icon this build doesn't
 * know, which is normal: a deployment can serve a format newer than the browser's cached bundle.
 */
export const GENERIC_OUTPUT: BlueprintOutput = {
  id: 'app',
  get noun() {
    return i18n._(GENERIC_OUTPUT_NOUN)
  },
  get plural() {
    return i18n._(GENERIC_OUTPUT_PLURAL)
  },
  icon: 'appWindow',
}

/** Resolve what to draw for a (possibly absent, possibly unrecognized) declared format. */
export function formatOf(output?: BlueprintOutput): BlueprintOutput {
  if (!output || !Object.hasOwn(FORMAT_ICONS, output.icon)) return GENERIC_OUTPUT
  return output
}

export function wireframeOf(output?: BlueprintOutput): FormatWireframe {
  return WIREFRAME_FOR_ICON[formatOf(output).icon]
}

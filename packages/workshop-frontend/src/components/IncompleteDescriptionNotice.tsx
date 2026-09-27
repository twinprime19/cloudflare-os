import { Warning } from '@phosphor-icons/react'
import type { ActionLogEntry } from '@gadgets/workshop-shared/api'
import { msg } from '@lingui/core/macro'
import { i18n } from '@gadgets/i18n'

/**
 * What an approver is told when the gatekeeper did not mark an action's description complete:
 * the text is a summary, truncated, or leaves out bytes that cannot be shown. A message
 * descriptor, not a resolved string: module scope evaluates before a locale is active, so the
 * render site resolves it with `i18n._()` against the locale active at render time.
 */
export const INCOMPLETE_DESCRIPTION_COPY = msg`This description is incomplete: part of what this action will send isn't shown.`

/**
 * True when a pending entry should carry the notice: an action its gatekeeper did not mark
 * complete. Hooks and observations send nothing, so they never do.
 */
export function isDescriptionIncomplete(entry: ActionLogEntry): boolean {
  return entry.type === 'action' && entry.description.descriptionIsComplete !== true
}

/**
 * Shown below a pending action's description on every approve surface (Activity review pane,
 * chat card, notifications popover) when the description is not complete. The action is still
 * submitted and approvable; the notice only tells the approver what they are not seeing.
 */
export function IncompleteDescriptionNotice({ id, className = '' }: { id?: string, className?: string }) {
  return (
    <div
      id={id}
      role="note"
      className={`flex items-start gap-2.5 rounded-2xl bg-kumo-tint px-3 py-2.5 ${className}`}
    >
      <div className="grid h-6 w-6 shrink-0 place-items-center text-kumo-warning">
        <Warning size={18} weight="duotone" />
      </div>
      <p className="m-0 text-[12px] leading-[18px] tracking-[-0.1px] text-kumo-default">
        {i18n._(INCOMPLETE_DESCRIPTION_COPY)}
      </p>
    </div>
  )
}

import { ShieldWarning } from '@phosphor-icons/react'
import { msg } from '@lingui/core/macro'
import { i18n } from '@gadgets/i18n'

/**
 * What an approver is told before approving an action in a workspace that has read restricted
 * data: the kernel does not check the action for that data, the approver does. A message
 * descriptor, not a resolved string: module scope evaluates before a locale is active, so the
 * render site resolves it with `i18n._()` against the locale active at render time.
 */
export const RESTRICTED_APPROVAL_COPY = msg`This workspace has read sensitive data. Read the full request below before approving — you are responsible for making sure it contains none of that data.`

/**
 * Shown above a pending action's description on every approve surface (Activity review pane,
 * chat card, notifications popover) while the workspace is restricted. The surfaces render the
 * description untruncated alongside it, since the notice asks the approver to read all of it.
 */
export function RestrictedApprovalNotice({ id, className = '' }: { id?: string, className?: string }) {
  return (
    <div
      id={id}
      role="note"
      className={`flex items-start gap-2.5 rounded-2xl bg-kumo-warning-tint px-3 py-2.5 ${className}`}
    >
      <div className="grid h-6 w-6 shrink-0 place-items-center text-kumo-warning">
        <ShieldWarning size={18} weight="duotone" />
      </div>
      <p className="m-0 text-[12px] leading-[18px] tracking-[-0.1px] text-kumo-default">
        {i18n._(RESTRICTED_APPROVAL_COPY)}
      </p>
    </div>
  )
}

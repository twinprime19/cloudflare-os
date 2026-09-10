/**
 * "Reconnecting…" pill for a fixed-height chrome strip (the workspace editor's top bar, the app
 * shell's top bar). Deliberately an inline chip rather than a full-width banner: a banner inserted
 * above the page reflows everything below it, so a blip that recovers on its own visibly jolts the
 * layout twice. `children` swaps the label for another connection state; with `onClick` the pill
 * is a button, for a state the user ends themselves (a paused workspace).
 */
import type { ReactNode } from 'react'
import { Trans } from '@lingui/react/macro'

const CHIP_CLASS =
  'text-xs text-kumo-warning px-2 py-0.5 rounded-full bg-kumo-warning-tint border border-kumo-warning/20'

export default function ReconnectingChip({ children, onClick }: {
  children?: ReactNode
  onClick?: () => void
}) {
  const label = children ?? <Trans>Reconnecting…</Trans>
  if (onClick) {
    return (
      <button type="button" onClick={onClick} className={`${CHIP_CLASS} cursor-pointer`}>
        {label}
      </button>
    )
  }
  return (
    <span role="status" className={CHIP_CLASS}>
      {label}
    </span>
  )
}

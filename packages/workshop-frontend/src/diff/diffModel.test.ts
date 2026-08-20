import { describe, expect, it } from 'vitest'
// `diffStatusLabel` resolves through the `t`/`msg` macro's global `@lingui/core` instance. This
// module has no component tree to mount a provider on, so importing `@gadgets/i18n` for its
// module-scope `i18n.loadAndActivate(...)` side effect is what activates it, the same as
// `main.tsx` does for the real app.
import '@gadgets/i18n'
import { diffStatusLabel, type DiffStatus } from './diffModel'

// The exact text `CodeDiffEditor.tsx`'s badge rendered before it was switched from the raw
// `DiffStatus` discriminant to `diffStatusLabel(model.status)`. `diffStatusLabel` must keep
// reproducing this set byte-identically -- these are also the literal discriminant values
// `buildDiffModel` assigns and other code compares against, so a mismatch here is a rendered-copy
// regression, not just a translation gap.
const TODAYS_BADGE_TEXT: Record<DiffStatus, string> = {
  Added: 'Added',
  Deleted: 'Deleted',
  Modified: 'Modified',
  Unchanged: 'Unchanged',
}

describe('diffStatusLabel', () => {
  it('reproduces the badge text the raw discriminant used to render, byte-identically', () => {
    for (const status of Object.keys(TODAYS_BADGE_TEXT) as DiffStatus[]) {
      expect(diffStatusLabel(status)).toBe(TODAYS_BADGE_TEXT[status])
    }
  })
})

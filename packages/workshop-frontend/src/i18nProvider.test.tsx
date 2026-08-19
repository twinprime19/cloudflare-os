// @vitest-environment jsdom
/* eslint-disable react/react-in-jsx-scope */

import { act, type ReactNode } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it } from 'vitest'
import { Trans } from '@lingui/react/macro'
import { AppI18nProvider, SOURCE_LOCALE, detectLocale, i18n } from '@gadgets/i18n'

// This suite lives in the app rather than in `@gadgets/i18n` because the pieces under test are the
// app's build wiring: the Babel macro pass that turns `<Trans>` into a runtime call, and the Lingui
// vite plugin that compiles `packages/i18n/locales/*.po` into an importable module. Both are
// configured in this package's `vite.config.ts`, which is also the config vitest runs under.
// Message resolution itself is covered by `i18nCatalog.test.ts`, which reproduces the production
// build's semantics rather than vitest's.

describe('AppI18nProvider', () => {
  let root: Root | undefined
  afterEach(() => {
    act(() => root?.unmount())
    root = undefined
  })

  function render(children: ReactNode): HTMLDivElement {
    const container = document.createElement('div')
    root = createRoot(container)
    act(() => {
      root!.render(<AppI18nProvider>{children}</AppI18nProvider>)
    })
    return container
  }

  it('renders its children on the first pass, with the source locale already active', () => {
    // `I18nProvider` renders `null` until a locale is active, so a provider that activated from an
    // effect would blank the whole app for a paint.
    expect(render(<Trans>Cancel</Trans>).textContent).toBe('Cancel')
    expect(i18n.locale).toBe(SOURCE_LOCALE)
  })

  it('resolves an interpolated message through the compiled catalog', () => {
    // The identifier's name is what the macro turns into the ICU placeholder, so this has to be
    // `actionLabel` to produce the same message -- and therefore the same id -- as the catalog
    // entry extracted from `AutoApproveConfirmDialog.tsx`.
    const actionLabel = 'Delete file'
    expect(render(<Trans>Always approve “{actionLabel}”?</Trans>).textContent)
      .toBe('Always approve “Delete file”?')
  })

  it('detects the source locale while English is the only catalog', () => {
    expect(detectLocale()).toBe(SOURCE_LOCALE)
  })
})

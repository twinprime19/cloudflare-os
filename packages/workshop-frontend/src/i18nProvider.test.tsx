// @vitest-environment jsdom
/* eslint-disable react/react-in-jsx-scope */

import { act, type ReactNode } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { Trans } from '@lingui/react/macro'
import { AppI18nProvider, SOURCE_LOCALE, detectLocale, i18n, type Locale } from '@gadgets/i18n'

// This suite lives in the app rather than in `@gadgets/i18n` because the pieces under test are the
// app's build wiring: the Babel macro pass that turns `<Trans>` into a runtime call, and the Lingui
// vite plugin that compiles `packages/i18n/locales/*.po` into an importable module. Both are
// configured in this package's `vite.config.ts`, which is also the config vitest runs under.
// Message resolution itself is covered by `i18nCatalog.test.ts`, which reproduces the production
// build's semantics rather than vitest's.

describe('AppI18nProvider', () => {
  let root: Root | undefined
  // jsdom builds one document for the whole file and the provider writes to its root element, so
  // the language claim outlives the container each test renders into and has to be put back by hand.
  const initialDocumentLanguage = document.documentElement.lang
  afterEach(() => {
    act(() => root?.unmount())
    root = undefined
    vi.unstubAllGlobals()
    // The instance under test is `@lingui/core`'s global singleton, so a test that activates
    // another locale would otherwise leave the next one resolving against that locale's catalog.
    if (i18n.locale !== SOURCE_LOCALE) i18n.activate(SOURCE_LOCALE)
    document.documentElement.lang = initialDocumentLanguage
  })

  function render(children: ReactNode): HTMLDivElement {
    const container = document.createElement('div')
    root = createRoot(container)
    act(() => {
      root!.render(<AppI18nProvider>{children}</AppI18nProvider>)
    })
    return container
  }

  /**
   * Renders under a locale the provider has to fetch a catalog for, and waits for the activation
   * that follows. The fetch is a dynamic import, which resolves through the module loader rather
   * than off the microtask queue, so flushing the render alone never observes it. Gives up after a
   * bounded wait and returns anyway, leaving the test's own assertions to report what happened.
   */
  async function renderLocale(locale: Locale, children: ReactNode): Promise<HTMLDivElement> {
    const container = document.createElement('div')
    root = createRoot(container)
    await act(async () => {
      root!.render(<AppI18nProvider locale={locale}>{children}</AppI18nProvider>)
    })
    for (let attempt = 0; attempt < 200 && i18n.locale !== locale; attempt++) {
      await act(async () => {
        await new Promise((resolve) => setTimeout(resolve, 5))
      })
    }
    // The loop exits on `i18n.locale`, which flips inside `activate()` -- before React has handled
    // the change event it emits. One more flush so the caller asserts against the re-rendered tree
    // rather than the one that was on screen a tick earlier.
    await act(async () => {})
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

  it('activates a non-source locale from its own catalog', async () => {
    // The same message the first test renders in English, so what this pins down is a locale
    // switch rather than the presence of a second catalog: one id, two locales, no component
    // change between them. "Hủy" is `vi.po`'s translation of "Cancel".
    const container = await renderLocale('vi', <Trans>Cancel</Trans>)
    expect(i18n.locale).toBe('vi')
    expect(container.textContent).toBe('Hủy')
  })

  it('claims the activated locale on the document element', async () => {
    // `index.html` ships `lang="en"` and nothing in a component updates it, so without this the
    // document would keep describing Vietnamese copy as English: a screen reader would read it with
    // English phonetics, and Chrome would offer to translate a page already in the reader's language.
    await renderLocale('vi', <Trans>Cancel</Trans>)
    expect(document.documentElement.lang).toBe('vi')
  })

  it('claims the source locale on the document element when it is the one that activates', () => {
    // Starting from a stale claim, because the case worth pinning down is the switch back: the
    // provider skips activating a locale that is already active, and the claim must move anyway.
    document.documentElement.lang = 'vi'
    render(<Trans>Cancel</Trans>)
    expect(document.documentElement.lang).toBe(SOURCE_LOCALE)
  })

  it('detects the source locale from the English preferences jsdom reports', () => {
    expect(detectLocale()).toBe(SOURCE_LOCALE)
  })

  it('detects a supported locale from the region variant a browser reports', () => {
    vi.stubGlobal('navigator', { languages: ['vi-VN', 'en'] })
    expect(detectLocale()).toBe('vi')
  })

  it('detects a supported locale a browser reports without a region', () => {
    vi.stubGlobal('navigator', { languages: ['vi'] })
    expect(detectLocale()).toBe('vi')
  })

  it('falls back to the source locale for a language with no catalog', () => {
    vi.stubGlobal('navigator', { languages: ['fr-FR'] })
    expect(detectLocale()).toBe(SOURCE_LOCALE)
  })

  it('falls back to the source locale outside a browser', () => {
    // What a caller in node sees -- vitest's own default environment, and anything server-side.
    vi.stubGlobal('navigator', undefined)
    expect(detectLocale()).toBe(SOURCE_LOCALE)
  })

  it('falls back to the source locale when no preferences are reported', () => {
    vi.stubGlobal('navigator', {})
    expect(detectLocale()).toBe(SOURCE_LOCALE)
  })
})

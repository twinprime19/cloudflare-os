// An ES import cannot pull in an ambient wildcard module declaration, which is the one thing a
// reference directive is for. The declaration has to reach the app's TypeScript program as well as
// this package's, and the app's tsconfig includes only the app's own `src` -- so it has to travel
// with the file that imports the catalog rather than be discovered from a config.
// oxlint-disable-next-line typescript/triple-slash-reference
/// <reference path="./po.d.ts" />
import { i18n, type Messages } from '@lingui/core'
import { I18nProvider } from '@lingui/react'
import { useEffect, type ReactNode } from 'react'
import { messages as sourceMessages } from '../locales/en.po'

/**
 * The configured Lingui instance. This is `@lingui/core`'s global singleton rather than one minted
 * by `setupI18n()`, because that is the instance the Babel macro injects into every compiled
 * `` t`...` `` call site. A second instance would activate a locale nothing reads.
 */
export { i18n }

/**
 * The language the source is written in, and the fallback for every other locale.
 *
 * Its catalog is imported statically above, not fetched, so it is compiled and active before any
 * component renders. That is load-bearing rather than an optimization: `descriptorFields: 'message'`
 * leaves each call site holding its message as *raw, uncompiled ICU*, and `@lingui/core` registers
 * its runtime ICU compiler only when `NODE_ENV !== 'production'`. So in a production build, a
 * message resolved from that inline fallback instead of from a compiled catalog renders its
 * placeholders literally -- `Always approve “{actionLabel}”?` would reach the user with the braces
 * still in it. A compiled catalog present from the first tick is what closes that window.
 */
export const SOURCE_LOCALE = 'en'

/**
 * Every locale with a catalog in `packages/i18n/locales`.
 *
 * Adding one is three lines, all of them inside this package and none of them in a component:
 * drop `locales/<tag>.po` beside `en.po`, add `'<tag>'` to `locales` in `lingui.config.ts`, and add
 * it here. Verified by activating a machine-generated catalog and asserting that already-written
 * components rendered it.
 *
 * The `lingui.config.ts` line is the one that is easy to miss and expensive to miss. Measured:
 * with the catalog present but the locale absent from `locales`, `@lingui/vite-plugin` compiles it
 * from the source language instead of from its own translations. The import succeeds, the module
 * exports the full set of message ids, `i18n.locale` reports the new locale, and every string on
 * screen is still English -- nothing thrown, nothing logged. `pnpm i18n:check` fails on that state
 * so it cannot reach a browser.
 */
export const SUPPORTED_LOCALES = [SOURCE_LOCALE, 'vi'] as const

/** A locale this app can activate. */
export type Locale = (typeof SUPPORTED_LOCALES)[number]

// Loaded and activated at import time, for the reason on SOURCE_LOCALE above. `@lingui/core` holds
// one catalog per locale, so activating another locale later does not evict this one.
i18n.loadAndActivate({ locale: SOURCE_LOCALE, messages: sourceMessages })

/** The shape `@lingui/vite-plugin` compiles a `.po` catalog into. */
interface CompiledCatalogModule {
  messages: Messages
}

/**
 * The locale to activate: the first entry of `navigator.languages` whose primary subtag names a
 * {@link SUPPORTED_LOCALES} entry. That list is the user's own preference order, so the first match
 * wins rather than the closest one. Matching on the primary subtag alone is what makes `vi-VN`
 * select `vi` -- there is one catalog per language here, and a region variant of a language we
 * translate is still that language. Falls back to {@link SOURCE_LOCALE}, which is also what a
 * caller outside a browser gets: there is no preference to read there.
 */
export function detectLocale(): Locale {
  if (typeof navigator === 'undefined') return SOURCE_LOCALE
  // `navigator.languages` is required of a browser but not of every host that defines a
  // `navigator`, and a browser may report it empty. Both cases mean "no preference", not a crash.
  for (const tag of navigator.languages ?? []) {
    const primarySubtag = tag.split('-')[0].toLowerCase()
    const supported = SUPPORTED_LOCALES.find((locale) => locale === primarySubtag)
    if (supported) return supported
  }
  return SOURCE_LOCALE
}

/**
 * Loads a non-source locale's compiled catalog. The specifier is a template literal so that no code
 * here names a locale: vite compiles the pattern into a glob over the directory, and
 * `@lingui/vite-plugin` turns each `.po` into a module exporting `messages`. Registering the locale
 * is still required -- see {@link SUPPORTED_LOCALES}.
 */
async function loadCatalog(locale: string): Promise<Messages> {
  const module = (await import(`../locales/${locale}.po`)) as Partial<CompiledCatalogModule>
  const messages = module.messages
  if (!messages || typeof messages !== 'object') {
    throw new Error(`The "${locale}" catalog module exports no messages.`)
  }
  return messages
}

/**
 * Points the document's own language claim at the locale that just became active.
 *
 * `index.html` ships `<html lang="en">` and no component owns the document element, so without this
 * the claim would stay English for the whole session no matter which catalog rendered. That claim is
 * read: a screen reader picks its pronunciation rules from it, so Vietnamese copy would be spoken
 * with English phonetics, and Chrome picks its translation prompt from it, so it would offer to
 * translate a page already served in the reader's language. It lives beside activation because the
 * claim is only true once a catalog is live -- announcing a language whose messages have not
 * resolved yet describes a screen that does not exist. Guarded for callers with no DOM, since this
 * module's provider and {@link detectLocale} both run under node in tests.
 */
function claimDocumentLanguage(locale: Locale): void {
  if (typeof document === 'undefined') return
  document.documentElement.lang = locale
}

/** Props for {@link AppI18nProvider}. */
export interface AppI18nProviderProps {
  /** Locale to activate. Defaults to {@link detectLocale}. */
  locale?: Locale
  children?: ReactNode
}

/**
 * The app's i18n root. Mount it above everything that renders copy -- including the error boundary,
 * whose fallback is copy of its own.
 */
export function AppI18nProvider({
  locale = detectLocale(),
  children,
}: AppI18nProviderProps): ReactNode {
  useEffect(() => {
    // The source locale is already loaded and active, so there is nothing to fetch and no window to
    // render through. The guard is for switching back to it from another locale; activating what is
    // already active would only emit a change event and force a pointless re-render.
    if (locale === SOURCE_LOCALE) {
      if (i18n.locale !== SOURCE_LOCALE) i18n.activate(SOURCE_LOCALE)
      // Unconditional, unlike the activation above: the source catalog is live either way by this
      // point, but the document may still be claiming the locale a previous render switched it to.
      claimDocumentLanguage(SOURCE_LOCALE)
      return
    }

    let cancelled = false
    loadCatalog(locale)
      .then((messages) => {
        // A locale change that lands while this fetch is in flight must win, so a catalog resolved
        // for the locale this effect started with is dropped rather than activated late.
        if (cancelled) return
        i18n.loadAndActivate({ locale, messages })
        claimDocumentLanguage(locale)
      })
      .catch((error: unknown) => {
        if (cancelled) return
        // The source catalog stays loaded and active, so a failed fetch degrades to correct English
        // rather than to raw ICU or a blank screen.
        console.error(`Could not load the "${locale}" message catalog.`, error)
      })
    return () => {
      cancelled = true
    }
  }, [locale])

  return <I18nProvider i18n={i18n}>{children}</I18nProvider>
}

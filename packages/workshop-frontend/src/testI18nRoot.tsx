// A React root for tests that mounts the app's i18n provider, the way `main.tsx` does.
//
// Components render copy through `<Trans>` and `` t`...` ``, both of which resolve against the
// Lingui instance `@gadgets/i18n` configures: `<Trans>` reads it from React context and throws
// without a provider, and `` t`...` `` reads the active locale off the global instance and throws
// if nothing activated one. In the app both are guaranteed by the single root in `main.tsx`, so a
// test that mounts a subtree has to supply the same root -- exactly as it would a router or an auth
// context. Wrapping here rather than at each `render` call keeps that a property of the root, which
// is what it is in production.
import { createRoot, type Root } from 'react-dom/client'
import { AppI18nProvider } from '@gadgets/i18n'
import type { ReactNode } from 'react'

/**
 * Creates a React root whose renders are wrapped in {@link AppI18nProvider}.
 *
 * Drop-in for `createRoot` from `react-dom/client`: it returns the same `Root` interface, so
 * `render` and `unmount` are called the same way.
 */
export function createI18nRoot(container: Element | DocumentFragment): Root {
  const root = createRoot(container)
  return {
    render: (children: ReactNode) => root.render(<AppI18nProvider>{children}</AppI18nProvider>),
    unmount: () => root.unmount(),
  }
}

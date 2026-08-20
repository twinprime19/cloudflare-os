import { defineConfig } from '@lingui/cli'
import { formatter } from '@lingui/format-po'

/**
 * The workspace's single i18n configuration. It lives here rather than beside the app because the
 * catalogs do: any package that adopts Lingui later adds an entry to `catalogs` instead of standing
 * up a second config. `<rootDir>` is this directory.
 *
 * `origins: false` drops the `#: file.tsx:line` comments Lingui writes by default. Two reasons, both
 * about this being a fork that syncs a very active upstream: the comments rewrite themselves on
 * every edit above a wrapped string, so `lingui extract` would produce a diff with no message
 * change in it, and every upstream sync that moves a line would manufacture a catalog conflict.
 * What is lost is the "where is this used" hint for translators; the message text is inline in the
 * source, so grepping the `msgid` recovers it exactly.
 */
/** The tree `lingui extract` scans. Named once because the exclusions have to repeat it. */
const APP_SOURCE = '<rootDir>/../workshop-frontend/src'

export default defineConfig({
  sourceLocale: 'en',
  locales: ['en', 'vi'],
  catalogs: [
    {
      path: '<rootDir>/locales/{locale}',
      include: [APP_SOURCE],
      // Test files hold fixture copy, and TanStack Router rewrites its generated route tree on
      // every build. Each pattern is anchored at `<rootDir>` rather than written as a bare
      // `**/*.test.tsx`: Lingui resolves `include` to absolute paths before handing the results to
      // `fs.globSync`, but passes `exclude` through untouched, so a relative pattern is compared
      // against an absolute path and silently matches nothing.
      exclude: [
        `${APP_SOURCE}/**/*.test.ts`,
        `${APP_SOURCE}/**/*.test.tsx`,
        `${APP_SOURCE}/**/*.gen.ts`,
        // Dead prototype UI: each reads `data/sample.ts` or `data/chat.ts` (already excluded from
        // the unlocalized-string guard, see the allowlist) and has no importer anywhere in a route
        // or a live component -- `ToolCallCard.tsx`'s only importer, `ChatMessage.tsx`, is itself
        // unreachable. Excluded rather than unwrapped: no code churn, no added upstream merge
        // surface, and the components arrive already localized if this UI is ever revived.
        `${APP_SOURCE}/components/chat/AppPreview.tsx`,
        `${APP_SOURCE}/components/chat/ConnectionConfigModal.tsx`,
        `${APP_SOURCE}/components/chat/DataTab.tsx`,
        `${APP_SOURCE}/components/chat/PermissionToast.tsx`,
        `${APP_SOURCE}/components/chat/ToolCallCard.tsx`,
      ],
    },
  ],
  format: formatter({ origins: false }),
})

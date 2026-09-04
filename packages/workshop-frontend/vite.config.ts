import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { defineConfig, loadEnv } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import tsconfigPaths from 'vite-tsconfig-paths'
import { TanStackRouterVite } from '@tanstack/router-plugin/vite'
import { lingui } from '@lingui/vite-plugin'
import { getConfig as getLinguiConfig } from '@lingui/conf'
import { vitestTask } from '@gadgets/scripts/vitest-task'

// `dist/` is this package's own build output, excluded from the inputs of the bundle and test
// tasks: vp declines to cache a task that reads a path it also wrote. Package-relative rather than
// workspace-wide -- `typed-storage` emits, and its `exports` resolves to `dist/index.js`, so a
// pattern matching every package's `dist` would drop a real input.
const frontendBundleTaskOptions = {
  dependsOn: ['clean:dist'],
  env: ['VITE_*'],
  input: [
    { auto: true },
    { pattern: '!dist/**', base: 'package' } as const,
    // Wrangler's scratch bundles are randomly named, and tracking reaches past the package that
    // owns the task, so any sibling that ran `wrangler dev` guarantees a miss here. Workspace-wide
    // for that reason; `build:app` and the shared `test` task exclude the same tree.
    { pattern: '!**/.wrangler/**', base: 'workspace' } as const,
  ],
  output: ['dist/**'],
}

const ownDist = { pattern: '!dist/**', base: 'package' } as const
const viteBuildCommand =
  `node --input-type=module -e "process.env.NODE_ENV='production'; await (await import('vite')).build()"`

// `@gadgets/i18n` owns `lingui.config.ts` and the catalogs beside it. Both Lingui plugins below
// otherwise discover that config by searching upwards from `process.cwd()`, which finds nothing
// from either of the two directories they actually run in: `vp run -r` loads this file from the
// workspace root to build its task graph, and the Babel pass runs with the cwd set to this package.
// So the config is resolved once, here, from a path relative to this file.
const linguiConfigCwd = resolve(dirname(fileURLToPath(import.meta.url)), '../i18n')
const linguiConfig = getLinguiConfig({ cwd: linguiConfigCwd })

const runConfig = {
  run: {
    tasks: {
      'clean:dist': {
        command: `node --input-type=module -e "import { rmSync } from 'node:fs'; rmSync('dist', { recursive: true, force: true })"`,
        cache: false,
      },
      /**
       * `build` is a task rather than a package.json script so `env` can declare the `VITE_*` flags
       * it reads: a cached `vp` run executes scripts in a clean environment, and the values would be
       * missing from the fingerprint besides. `VITE_CF_ACCESS_MODE` is inlined into the bundle
       * (`src/useAuth.ts`) and `VITE_FRONTEND_ERROR_REPORTING` selects hidden source maps below, so
       * replaying a bundle built under different values is wrong rather than merely stale.
       *
       * Separate commands rather than one `&&` string so each is a cache entry of its own; `env` is
       * task-wide, so changing a flag re-runs all three anyway. `tsconfig.vite.json` is the
       * config-file pass (`vite.config.ts` and the scripts it imports), which the app's own
       * `tsconfig.json` excludes.
       *
       * Production mode is set before Vite is imported because Vite snapshots whether `NODE_ENV`
       * was present before loading this config. The Node launcher is shell-neutral.
       */
      build: {
        command: ['tsc', 'tsc -p tsconfig.vite.json', viteBuildCommand],
        ...frontendBundleTaskOptions,
      },
      'build:assets': {
        command: viteBuildCommand,
        ...frontendBundleTaskOptions,
      },
      test: vitestTask('vitest run', [ownDist]),
    },
  },
}

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd())
  const backendHost = env.VITE_BACKEND_HOST?.trim() || 'localhost:8787'
  const frontendErrorReporting = env.VITE_FRONTEND_ERROR_REPORTING === 'true'
  return {
    // Spread, not a literal `run: {...}`: `run` is Vite+'s field and vite's own `defineConfig` has
    // no such property, but the excess-property check doesn't reach spreads.
    ...runConfig,
    plugins: [
      TanStackRouterVite({ target: 'react', autoCodeSplitting: true }),
      // The Lingui macros (`<Trans>`, `` t`...` ``) are a compile-time transform, so they ride
      // plugin-react's existing Babel pass rather than adding a second one. This is scoped to this
      // package and leaves the workspace-wide esbuild transform untouched.
      //
      // `descriptorFields: 'message'` rather than the default `'auto'`: `'auto'` compiles a
      // production build down to the message id alone, which renders as a bare hash unless a
      // compiled catalog for the active locale is loaded and still matches. English is the source
      // language and lives inline here, so keeping the message beside the id makes the source text
      // its own fallback -- a missed `lingui extract` can then never blank out the UI.
      react({
        babel: {
          plugins: [
            ['@lingui/babel-plugin-lingui-macro', { descriptorFields: 'message', linguiConfig }],
          ],
        },
      }),
      // Compiles `.po` catalogs on import.
      lingui({ cwd: linguiConfigCwd }),
      tailwindcss(),
      tsconfigPaths(),
    ],
    server: {
      port: 3000,
      host: true,
      proxy: {
        '/api/client-errors': `http://${backendHost}`,
        '/blueprint-screenshot': `http://${backendHost}`,
        '/api/site-logo': `http://${backendHost}`,
      },
    },
    build: {
      // Production reporting uploads these separately; hidden maps never reveal a map URL to users.
      sourcemap: frontendErrorReporting ? 'hidden' : false,
    },
  }
})

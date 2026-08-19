/**
 * `@lingui/vite-plugin` turns a `.po` catalog into a module exporting the compiled messages. The
 * declaration is referenced from `index.tsx` rather than left for tsconfig to discover, so it
 * reaches every program that pulls this package in -- the app's `tsc` pass sees this package as
 * source, and its own tsconfig includes only its own `src`.
 */
declare module "*.po" {
  const messages: import("@lingui/core").Messages
  export { messages }
}

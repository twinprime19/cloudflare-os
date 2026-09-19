// React's view of the deliberate socket pause.
//
// `connectionPause.ts` is deliberately React-free -- `main.tsx` owns the socket outside the
// component tree -- so anything that renders the pause mirrors the flag through this hook. Both
// chips read it: the workspace editor's header and the app shell's top bar. The shell needs its own
// because a pause the server starts can land on any route -- a page that holds an Overseer stub
// open, a Share modal for instance, parks with no editor mounted -- and without a chip there the
// only way out is a reload.

import { useSyncExternalStore } from 'react'
import { isConnectionPaused, subscribeConnectionPause } from './connectionPause'

/** True while the connection is parked; re-renders on every pause and resume. */
export function useConnectionPaused(): boolean {
  return useSyncExternalStore(subscribeConnectionPause, isConnectionPaused)
}

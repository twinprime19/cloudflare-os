// Stand-in for the `cloudflare:workers` module, which only exists inside workerd.
//
// The gatekeeper module declares a Durable Object and RPC entrypoints, so it cannot be imported
// under plain vitest without it. Only the base classes are provided, and only so that `import` and
// `extends` resolve -- the tests drive the classes with their own fake context objects, and
// anything that genuinely needs the runtime belongs in a Workers-pool test, not here.

export class DurableObject<E = unknown, P = unknown> {
  constructor(readonly ctx: unknown, readonly env: E, readonly props?: P) {}
}

export class RpcTarget {}

export class WorkerEntrypoint<E = unknown, P = unknown> {
  constructor(readonly ctx: unknown, readonly env: E, readonly props?: P) {}
}

export class RpcStub<T> {
  constructor(readonly target: T) {}
}

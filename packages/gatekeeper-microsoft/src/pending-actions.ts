// Queued-but-undecided actions, in the gatekeeper Durable Object's synchronous KV.
//
// One counter plus one record per action: the store hands out the sequential id the approval queue
// identifies an action by, and holds the caller's intent until applyAction()/rejectAction() removes
// it. It is deliberately unaware of what an action means — each gatekeeper supplies its own action
// type and rebuilds the request at apply time.
//
// Shared by the Outlook mailbox and SharePoint list gatekeepers, which need the same bookkeeping.

/** Ceiling on queued-but-undecided actions, matching the Gmail gatekeeper. */
export const MAX_PENDING_ACTIONS = 100;

export class PendingActionStore<Action> {
  #kv: DurableObjectStorage["kv"];

  constructor(kv: DurableObjectStorage["kv"]) {
    this.#kv = kv;
  }

  #actionKey(id: number): string {
    return `pending:action:${id}`;
  }

  submit(action: Action): number {
    let id = this.#kv.get<number>("pending:nextActionId") ?? 1;
    this.#kv.put("pending:nextActionId", id + 1);
    this.#kv.put(this.#actionKey(id), action);
    return id;
  }

  get(id: number): Action | undefined {
    return this.#kv.get<Action>(this.#actionKey(id));
  }

  list(): {id: number, action: Action}[] {
    return [...this.#kv.list<Action>({prefix: "pending:action:"})]
        .map(([key, action]) => ({id: Number(key.slice("pending:action:".length)), action}))
        .filter(({id}) => Number.isFinite(id))
        .toSorted((a, b) => a.id - b.id);
  }

  remove(id: number): void {
    this.#kv.delete(this.#actionKey(id));
  }
}

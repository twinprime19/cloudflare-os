// Which calls renew a workspace's client-activity lease (OverseerImpl.noteClientRpc). What decides
// is provenance carried by the instance, never the class: the agent, hook deliveries and binding
// loopbacks mint the very same capability classes a browser session is handed. So the cases that
// discriminate are two instances of one class -- one minted for a browser, one minted internally --
// plus the approval queue a gatekeeper session holds, which is never wrapped at all: a gatekeeper
// answering its session is not a user being here, and taking it for one is what would keep a
// workspace resident with nobody in front of it.
//
// The lease's clock and its expiry run against a real Durable Object in
// __integration__/session-lease.test.ts; what is asserted here is only who renews it.

import { describe, expect, it, vi } from "vitest";
import { env } from "cloudflare:workers";
import { runInDurableObject } from "cloudflare:test";
import type { OverseerDurableObject } from "../src/overseer.js";
import { openFakeOverseer } from "./fixtures.js";

declare module "cloudflare:workers" {
  interface ProvidedEnv {
    TEST_OVERSEER: DurableObjectNamespace<OverseerDurableObject>;
  }
}

const CONNECTION_SPEC = {
  type: "gatekeeper" as const,
  vendorId: "testvendor",
  resourceUrl: "https://example.com/new",
  typeUrlPattern: "https://*",
};

// The impl that the Durable Object and the forged client interface each keep is TypeScript-private,
// and the lease renewal lives on it. Reaching it is the same visibility escape the lease's
// integration tests make; typed as `any` because what is reached is the internal surface, not a
// contract.
function asImpl(holder: object): any {
  return (holder as { impl: unknown }).impl;
}

describe("client-activity lease provenance", () => {
  it("renews only for a capability minted for a browser session", async () => {
    const stub = env.TEST_OVERSEER.getByName("session-lease-provenance");
    await runInDurableObject(stub, async (instance: OverseerDurableObject) => {
      const impl = asImpl(instance);
      // Every renewal this workspace is told about, whoever makes it.
      const renewals = vi.fn();
      impl.noteClientRpc = renewals;

      // A connection facet that answers what adding one and opening a session on it need, and
      // keeps the approval queue the session is opened with.
      let approvalQueue: { getGitCache(): Promise<unknown> } | undefined;
      const getGatekeeperFacet = () => ({
        describe: async () => ({ title: "Test", url: "https://example.com/new" }),
        startSession: (queue: typeof approvalQueue) => { approvalQueue = queue; return {}; },
      });
      impl.getGatekeeperFacet = getGatekeeperFacet;

      // The internal mint: the capability a binding loopback or the agent works through. No browser
      // is behind it, so neither its own calls nor the approval queue it opens a session with may
      // pass for someone being here.
      const internal = await impl.addGatekeeper({}, CONNECTION_SPEC);
      const gatekeeperId: number = await internal.getId();
      await internal.openSession();
      await approvalQueue!.getGitCache();
      expect(renewals).not.toHaveBeenCalled();

      // The same class, minted for a browser session -- over this workspace's own connection table,
      // so the two capabilities differ in nothing but their provenance.
      const client = await openFakeOverseer(
          { gatekeepers: impl.storage.gatekeepers }, { impl: { getGatekeeperFacet } });
      const browserRenewals = asImpl(client).noteClientRpc;
      // open() is itself a browser call; what follows it is the comparison.
      browserRenewals.mockClear();

      const minted = await client.getGatekeeperById(gatekeeperId);
      expect(browserRenewals).toHaveBeenCalledTimes(1);
      expect(await minted.getId()).toBe(gatekeeperId);
      expect(browserRenewals).toHaveBeenCalledTimes(2);

      // And that traffic was this workspace's, not the one the internal capability belongs to.
      expect(renewals).not.toHaveBeenCalled();
    });
  });
});

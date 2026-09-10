import { createObservabilityContext } from "@gadgets/backend-utils/observability-context";

/** Observability fields emitted by the Microsoft gatekeeper. */
export type MicrosoftObservabilityFields = {
  vendorId: string;
  /** `urlPattern` of the grantable resource an event is about. */
  resource: string;
  /** OAuth scopes a consent round trip asked for and did not come back with. Names only. */
  missingScopes: string[];
  /** OAuth scopes the grant does cover, so the delta can be read without a second lookup. */
  grantedScopes: string[];
};

/** Ambient observability fields for one Microsoft gatekeeper operation. */
export const obsContext = createObservabilityContext<MicrosoftObservabilityFields>();

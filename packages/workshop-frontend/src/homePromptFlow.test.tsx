// @vitest-environment jsdom
/* eslint-disable react/react-in-jsx-scope */

import { act } from "react";
import { type Root } from "react-dom/client";
import { createI18nRoot } from "./testI18nRoot";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { BlueprintBindingAssignment, OutputFormatOffer } from "@gadgets/workshop-shared/api";

const DOC_FORMAT: OutputFormatOffer = {
  blueprintId: "format.document",
  output: { id: "document", noun: "Doc", plural: "Docs", icon: "fileText" },
  description: "Write documents",
  requiresSetup: false,
};

const SLIDES_FORMAT: OutputFormatOffer = {
  blueprintId: "format.slides",
  output: { id: "slides", noun: "Contoso Slides", plural: "Contoso Slides", icon: "presentation" },
  description: "",
  requiresSetup: false,
};

const testState = vi.hoisted(() => {
  const listModels = vi.fn<() => Promise<never[]>>(async () => []);
  const newGadget = vi.fn<() => never>();
  const listOutputFormats = vi.fn<() => Promise<OutputFormatOffer[]>>(async () => []);
  const newGadgetFromBlueprint = vi.fn<
    (blueprintId: string, bindings: Record<string, BlueprintBindingAssignment>) => Promise<unknown>
  >();
  // A new object each time so useOutputFormats's module cache (keyed on the api stub) does not
  // replay a previous test's list across this file.
  const makeApi = () => ({
    listModels,
    newGadget,
    listOutputFormats,
    newGadgetFromBlueprint,
  });
  return {
    addToast: vi.fn<(toast: unknown) => void>(),
    authenticatedApi: makeApi(),
    makeApi,
    currentUser: { id: "user-a", name: "User A" },
    listModels,
    listOutputFormats,
    navigate: vi.fn<(options: unknown) => void>(),
    newGadget,
    newGadgetFromBlueprint,
    seeds: [] as Array<{ text?: string; nonce?: number }>,
    draftStorageKeys: [] as Array<string | undefined>,
  };
});

vi.mock("@tanstack/react-router", async (importOriginal) => ({
  ...(await importOriginal<typeof import("@tanstack/react-router")>()),
  useNavigate: () => testState.navigate,
}));

vi.mock("@cloudflare/kumo", () => ({
  useKumoToastManager: () => ({ add: testState.addToast }),
}));

vi.mock("./AuthContext", () => ({
  useAuthenticatedApi: () => ({
    authenticatedApi: testState.authenticatedApi,
    currentUser: testState.currentUser,
  }),
}));

vi.mock("./features/chat/composer/ChatComposer", () => ({
  ChatComposer: ({ seedText, seedNonce, draftStorageKey }: {
    seedText?: string;
    seedNonce?: number;
    draftStorageKey?: string;
  }) => {
    testState.seeds.push({ text: seedText, nonce: seedNonce });
    testState.draftStorageKeys.push(draftStorageKey);
    return <textarea aria-label="Prompt" readOnly value={seedText ?? ""} />;
  },
}));

vi.mock("./components/MeshBackground", () => ({ default: () => null }));
vi.mock("./components/AppShell/HomeTaskSuggestions", () => ({ default: () => null }));
vi.mock("./useDocumentTitle", () => ({ useDocumentTitle: () => {} }));

import { HomePageContent } from "./routes/index";

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

describe("Home prompt route flow", () => {
  let container: HTMLDivElement | undefined;
  let root: Root | undefined;

  afterEach(async () => {
    await act(async () => root?.unmount());
    container?.remove();
    localStorage.clear();
    testState.seeds.length = 0;
    testState.draftStorageKeys.length = 0;
    vi.clearAllMocks();
    testState.listOutputFormats.mockResolvedValue([]);
    testState.authenticatedApi = testState.makeApi();
  });

  it("seeds the composer once, clears route state, and does not create a workspace", async () => {
    container = document.createElement("div");
    document.body.append(container);
    root = createI18nRoot(container);
    await act(async () => root!.render(<HomePageContent prompt="Create a daily brief." />));

    expect(container.querySelector<HTMLTextAreaElement>('[aria-label="Prompt"]')?.value).toBe(
      "Create a daily brief.",
    );
    expect(Math.max(...testState.seeds.map(({ nonce }) => nonce ?? 0))).toBe(1);
    expect(testState.navigate).toHaveBeenCalledWith({ to: "/", search: {}, replace: true });
    expect(testState.newGadget).not.toHaveBeenCalled();
    expect(testState.draftStorageKeys).toContain("gadgets:composer-draft:v1:user-a:home");
  });

  it("offers each promoted format as a New shortcut that stamps a workspace with no chat", async () => {
    testState.listOutputFormats.mockResolvedValue([DOC_FORMAT, SLIDES_FORMAT]);
    const overseer = {
      getMetadata: vi.fn<() => Promise<{ id: string }>>(async () => ({ id: "ws-doc" })),
      [Symbol.dispose]: vi.fn<() => void>(),
    };
    testState.newGadgetFromBlueprint.mockResolvedValue(overseer);

    container = document.createElement("div");
    document.body.append(container);
    root = createI18nRoot(container);
    await act(async () => root!.render(<HomePageContent />));
    await act(async () => { await Promise.resolve(); });

    const buttons = [...container.querySelectorAll("button")];
    expect(buttons.map((button) => button.textContent)).toEqual(
      expect.arrayContaining(["New Doc", "New Contoso Slides"]),
    );

    const newDoc = buttons.find((button) => button.textContent === "New Doc");
    expect(newDoc).toBeDefined();
    await act(async () => { newDoc!.click(); });

    expect(testState.newGadgetFromBlueprint).toHaveBeenCalledWith("format.document", {});
    expect(testState.newGadget).not.toHaveBeenCalled();
    expect(testState.navigate).toHaveBeenCalledWith({
      to: "/workspace/$id",
      params: { id: "ws-doc" },
    });
    expect(overseer[Symbol.dispose]).toHaveBeenCalled();
  });
});

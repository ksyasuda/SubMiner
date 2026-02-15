import test from "node:test";
import assert from "node:assert/strict";
import { AppReadyRuntimeDeps, runAppReadyRuntimeService } from "./startup-service";

function makeDeps(overrides: Partial<AppReadyRuntimeDeps> = {}) {
  const calls: string[] = [];
  const deps: AppReadyRuntimeDeps = {
    loadSubtitlePosition: () => calls.push("loadSubtitlePosition"),
    resolveKeybindings: () => calls.push("resolveKeybindings"),
    createMpvClient: () => calls.push("createMpvClient"),
    reloadConfig: () => calls.push("reloadConfig"),
    getResolvedConfig: () => ({ websocket: { enabled: "auto" }, secondarySub: {} }),
    getConfigWarnings: () => [],
    logConfigWarning: () => calls.push("logConfigWarning"),
    initRuntimeOptionsManager: () => calls.push("initRuntimeOptionsManager"),
    setSecondarySubMode: (mode) => calls.push(`setSecondarySubMode:${mode}`),
    defaultSecondarySubMode: "hover",
    defaultWebsocketPort: 9001,
    hasMpvWebsocketPlugin: () => true,
    startSubtitleWebsocket: (port) => calls.push(`startSubtitleWebsocket:${port}`),
    log: (message) => calls.push(`log:${message}`),
    createMecabTokenizerAndCheck: async () => {
      calls.push("createMecabTokenizerAndCheck");
    },
    createSubtitleTimingTracker: () => calls.push("createSubtitleTimingTracker"),
    loadYomitanExtension: async () => {
      calls.push("loadYomitanExtension");
    },
    texthookerOnlyMode: false,
    shouldAutoInitializeOverlayRuntimeFromConfig: () => true,
    initializeOverlayRuntime: () => calls.push("initializeOverlayRuntime"),
    handleInitialArgs: () => calls.push("handleInitialArgs"),
    ...overrides,
  };
  return { deps, calls };
}

test("runAppReadyRuntimeService starts websocket in auto mode when plugin missing", async () => {
  const { deps, calls } = makeDeps({
    hasMpvWebsocketPlugin: () => false,
  });
  await runAppReadyRuntimeService(deps);
  assert.ok(calls.includes("startSubtitleWebsocket:9001"));
  assert.ok(calls.includes("initializeOverlayRuntime"));
});

test("runAppReadyRuntimeService logs defer message when overlay not auto-started", async () => {
  const { deps, calls } = makeDeps({
    shouldAutoInitializeOverlayRuntimeFromConfig: () => false,
  });
  await runAppReadyRuntimeService(deps);
  assert.ok(
    calls.includes(
      "log:Overlay runtime deferred: waiting for explicit overlay command.",
    ),
  );
});

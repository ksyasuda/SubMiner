import test from "node:test";
import assert from "node:assert/strict";
import {
  createStartupAppReadyDepsRuntimeService,
  createStartupAppShutdownDepsRuntimeService,
} from "./startup-lifecycle-runtime-deps-service";

test("createStartupAppReadyDepsRuntimeService preserves runtime deps behavior", async () => {
  const calls: string[] = [];
  const deps = createStartupAppReadyDepsRuntimeService({
    loadSubtitlePosition: () => calls.push("loadSubtitlePosition"),
    resolveKeybindings: () => calls.push("resolveKeybindings"),
    createMpvClient: () => calls.push("createMpvClient"),
    reloadConfig: () => calls.push("reloadConfig"),
    getResolvedConfig: () => ({
      secondarySub: { defaultMode: "hover" },
      websocket: { enabled: "auto", port: 1234 },
    }),
    getConfigWarnings: () => [],
    logConfigWarning: () => {},
    initRuntimeOptionsManager: () => calls.push("initRuntimeOptionsManager"),
    setSecondarySubMode: () => calls.push("setSecondarySubMode"),
    defaultSecondarySubMode: "hover",
    defaultWebsocketPort: 8765,
    hasMpvWebsocketPlugin: () => true,
    startSubtitleWebsocket: () => calls.push("startSubtitleWebsocket"),
    log: () => calls.push("log"),
    createMecabTokenizerAndCheck: async () => {
      calls.push("createMecab");
    },
    createSubtitleTimingTracker: () => calls.push("createSubtitleTimingTracker"),
    loadYomitanExtension: async () => {
      calls.push("loadYomitan");
    },
    texthookerOnlyMode: false,
    shouldAutoInitializeOverlayRuntimeFromConfig: () => false,
    initializeOverlayRuntime: () => calls.push("initOverlayRuntime"),
    handleInitialArgs: () => calls.push("handleInitialArgs"),
  });

  deps.loadSubtitlePosition();
  await deps.createMecabTokenizerAndCheck();
  deps.handleInitialArgs();

  assert.equal(deps.defaultWebsocketPort, 8765);
  assert.equal(deps.defaultSecondarySubMode, "hover");
  assert.deepEqual(calls, ["loadSubtitlePosition", "createMecab", "handleInitialArgs"]);
});

test("createStartupAppShutdownDepsRuntimeService preserves shutdown handlers", () => {
  const calls: string[] = [];
  const deps = createStartupAppShutdownDepsRuntimeService({
    unregisterAllGlobalShortcuts: () => calls.push("unregisterAllGlobalShortcuts"),
    stopSubtitleWebsocket: () => calls.push("stopSubtitleWebsocket"),
    stopTexthookerService: () => calls.push("stopTexthookerService"),
    destroyYomitanParserWindow: () => calls.push("destroyYomitanParserWindow"),
    clearYomitanParserPromises: () => calls.push("clearYomitanParserPromises"),
    stopWindowTracker: () => calls.push("stopWindowTracker"),
    destroyMpvSocket: () => calls.push("destroyMpvSocket"),
    clearReconnectTimer: () => calls.push("clearReconnectTimer"),
    destroySubtitleTimingTracker: () => calls.push("destroySubtitleTimingTracker"),
    destroyAnkiIntegration: () => calls.push("destroyAnkiIntegration"),
  });

  deps.stopSubtitleWebsocket();
  deps.clearReconnectTimer();
  deps.destroyAnkiIntegration();

  assert.deepEqual(calls, [
    "stopSubtitleWebsocket",
    "clearReconnectTimer",
    "destroyAnkiIntegration",
  ]);
});

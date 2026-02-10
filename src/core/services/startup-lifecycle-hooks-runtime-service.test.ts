import test from "node:test";
import assert from "node:assert/strict";
import { createStartupLifecycleHooksRuntimeService } from "./startup-lifecycle-hooks-runtime-service";

test("createStartupLifecycleHooksRuntimeService wires app-ready and app-shutdown handlers", async () => {
  const calls: string[] = [];
  const hooks = createStartupLifecycleHooksRuntimeService({
    appReadyDeps: {
      loadSubtitlePosition: () => calls.push("loadSubtitlePosition"),
      resolveKeybindings: () => calls.push("resolveKeybindings"),
      createMpvClient: () => calls.push("createMpvClient"),
      reloadConfig: () => calls.push("reloadConfig"),
      getResolvedConfig: () => ({
        secondarySub: { defaultMode: "hover" },
        websocket: { enabled: "auto", port: 1234 },
      }),
      getConfigWarnings: () => [],
      logConfigWarning: () => calls.push("logConfigWarning"),
      initRuntimeOptionsManager: () => calls.push("initRuntimeOptionsManager"),
      setSecondarySubMode: () => calls.push("setSecondarySubMode"),
      defaultSecondarySubMode: "hover",
      defaultWebsocketPort: 8765,
      hasMpvWebsocketPlugin: () => true,
      startSubtitleWebsocket: () => calls.push("startSubtitleWebsocket"),
      log: () => calls.push("log"),
      createMecabTokenizerAndCheck: async () => {
        calls.push("createMecabTokenizerAndCheck");
      },
      createSubtitleTimingTracker: () => calls.push("createSubtitleTimingTracker"),
      loadYomitanExtension: async () => {
        calls.push("loadYomitanExtension");
      },
      texthookerOnlyMode: false,
      shouldAutoInitializeOverlayRuntimeFromConfig: () => false,
      initializeOverlayRuntime: () => calls.push("initializeOverlayRuntime"),
      handleInitialArgs: () => calls.push("handleInitialArgs"),
    },
    appShutdownDeps: {
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
    },
    shouldRestoreWindowsOnActivate: () => true,
    restoreWindowsOnActivate: () => calls.push("restoreWindowsOnActivate"),
  });

  await hooks.onReady();
  hooks.onWillQuitCleanup();
  assert.equal(hooks.shouldRestoreWindowsOnActivate(), true);
  hooks.restoreWindowsOnActivate();

  assert.ok(calls.includes("loadSubtitlePosition"));
  assert.ok(calls.includes("handleInitialArgs"));
  assert.ok(calls.includes("destroyAnkiIntegration"));
  assert.ok(calls.includes("restoreWindowsOnActivate"));
});

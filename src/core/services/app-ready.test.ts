import test from "node:test";
import assert from "node:assert/strict";
import { AppReadyRuntimeDeps, runAppReadyRuntime } from "./startup";

function makeDeps(overrides: Partial<AppReadyRuntimeDeps> = {}) {
  const calls: string[] = [];
  const deps: AppReadyRuntimeDeps = {
    loadSubtitlePosition: () => calls.push("loadSubtitlePosition"),
    resolveKeybindings: () => calls.push("resolveKeybindings"),
    createMpvClient: () => calls.push("createMpvClient"),
    reloadConfig: () => calls.push("reloadConfig"),
    getResolvedConfig: () => ({
      websocket: { enabled: "auto" },
      secondarySub: {},
    }),
    getConfigWarnings: () => [],
    logConfigWarning: () => calls.push("logConfigWarning"),
    setLogLevel: (level, source) =>
      calls.push(`setLogLevel:${level}:${source}`),
    initRuntimeOptionsManager: () => calls.push("initRuntimeOptionsManager"),
    setSecondarySubMode: (mode) => calls.push(`setSecondarySubMode:${mode}`),
    defaultSecondarySubMode: "hover",
    defaultWebsocketPort: 9001,
    hasMpvWebsocketPlugin: () => true,
    startSubtitleWebsocket: (port) =>
      calls.push(`startSubtitleWebsocket:${port}`),
    log: (message) => calls.push(`log:${message}`),
    createMecabTokenizerAndCheck: async () => {
      calls.push("createMecabTokenizerAndCheck");
    },
    createSubtitleTimingTracker: () =>
      calls.push("createSubtitleTimingTracker"),
    createImmersionTracker: () => calls.push("createImmersionTracker"),
    startJellyfinRemoteSession: async () => {
      calls.push("startJellyfinRemoteSession");
    },
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

test("runAppReadyRuntime starts websocket in auto mode when plugin missing", async () => {
  const { deps, calls } = makeDeps({
    hasMpvWebsocketPlugin: () => false,
  });
  await runAppReadyRuntime(deps);
  assert.ok(calls.includes("startSubtitleWebsocket:9001"));
  assert.ok(calls.includes("initializeOverlayRuntime"));
  assert.ok(calls.includes("createImmersionTracker"));
  assert.ok(calls.includes("startJellyfinRemoteSession"));
  assert.ok(
    calls.includes("log:Runtime ready: invoking createImmersionTracker."),
  );
});

test("runAppReadyRuntime skips Jellyfin remote startup when dependency is not wired", async () => {
  const { deps, calls } = makeDeps({
    startJellyfinRemoteSession: undefined,
  });

  await runAppReadyRuntime(deps);

  assert.equal(calls.includes("startJellyfinRemoteSession"), false);
  assert.ok(calls.includes("createMecabTokenizerAndCheck"));
  assert.ok(calls.includes("createMpvClient"));
  assert.ok(calls.includes("createSubtitleTimingTracker"));
  assert.ok(calls.includes("handleInitialArgs"));
  assert.ok(
    calls.includes("initializeOverlayRuntime") ||
      calls.includes(
        "log:Overlay runtime deferred: waiting for explicit overlay command.",
      ),
  );
});

test("runAppReadyRuntime logs when createImmersionTracker dependency is missing", async () => {
  const { deps, calls } = makeDeps({
    createImmersionTracker: undefined,
  });
  await runAppReadyRuntime(deps);
  assert.ok(
    calls.includes(
      "log:Runtime ready: createImmersionTracker dependency is missing.",
    ),
  );
});

test("runAppReadyRuntime logs and continues when createImmersionTracker throws", async () => {
  const { deps, calls } = makeDeps({
    createImmersionTracker: () => {
      calls.push("createImmersionTracker");
      throw new Error("immersion init failed");
    },
  });
  await runAppReadyRuntime(deps);
  assert.ok(calls.includes("createImmersionTracker"));
  assert.ok(
    calls.includes(
      "log:Runtime ready: createImmersionTracker failed: immersion init failed",
    ),
  );
  assert.ok(calls.includes("initializeOverlayRuntime"));
  assert.ok(calls.includes("handleInitialArgs"));
});

test("runAppReadyRuntime logs defer message when overlay not auto-started", async () => {
  const { deps, calls } = makeDeps({
    shouldAutoInitializeOverlayRuntimeFromConfig: () => false,
  });
  await runAppReadyRuntime(deps);
  assert.ok(
    calls.includes(
      "log:Overlay runtime deferred: waiting for explicit overlay command.",
    ),
  );
});

test("runAppReadyRuntime applies config logging level during app-ready", async () => {
  const { deps, calls } = makeDeps({
    getResolvedConfig: () => ({
      websocket: { enabled: "auto" },
      secondarySub: {},
      logging: { level: "warn" },
    }),
  });
  await runAppReadyRuntime(deps);
  assert.ok(calls.includes("setLogLevel:warn:config"));
});

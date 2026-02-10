import test from "node:test";
import assert from "node:assert/strict";
import { createIpcDepsRuntimeService } from "./ipc-deps-runtime-service";

test("createIpcDepsRuntimeService maps window and mecab helpers", async () => {
  let ignoreMouse: { ignore: boolean; forward?: boolean } | null = null;
  let toggledDevTools = 0;
  let mecabEnabled: boolean | null = null;

  const visibleWindow = {
    isDestroyed: () => false,
    setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
      ignoreMouse = { ignore, forward: options?.forward };
    },
    webContents: {
      toggleDevTools: () => {
        toggledDevTools += 1;
      },
    },
  };

  const deps = createIpcDepsRuntimeService({
    getInvisibleWindow: () => visibleWindow,
    getMainWindow: () => visibleWindow,
    getVisibleOverlayVisibility: () => true,
    getInvisibleOverlayVisibility: () => false,
    onOverlayModalClosed: () => {},
    openYomitanSettings: () => {},
    quitApp: () => {},
    toggleVisibleOverlay: () => {},
    tokenizeCurrentSubtitle: async () => ({ text: "x" }),
    getCurrentSubtitleAss: () => "ass",
    getMpvSubtitleRenderMetrics: () => ({ subPos: 100 }),
    getSubtitlePosition: () => ({ x: 1, y: 2 }),
    getSubtitleStyle: () => null,
    saveSubtitlePosition: () => {},
    getMecabTokenizer: () => ({
      getStatus: () => ({ available: true, enabled: true, path: "/usr/bin/mecab" }),
      setEnabled: (enabled: boolean) => {
        mecabEnabled = enabled;
      },
    }),
    handleMpvCommand: () => {},
    getKeybindings: () => ({ copySubtitle: ["C"] }),
    getSecondarySubMode: () => "hidden",
    getMpvClient: () => ({ currentSecondarySubText: "secondary" }),
    runSubsyncManual: async () => ({ ok: true }),
    getAnkiConnectStatus: () => true,
    getRuntimeOptions: () => ({ values: {} }),
    setRuntimeOption: () => ({ ok: true }),
    cycleRuntimeOption: () => ({ ok: true }),
  });

  deps.setInvisibleIgnoreMouseEvents(true, { forward: true });
  deps.toggleDevTools();
  deps.setMecabEnabled(false);

  assert.deepEqual(ignoreMouse, { ignore: true, forward: true });
  assert.equal(toggledDevTools, 1);
  assert.equal(mecabEnabled, false);
  assert.deepEqual(deps.getMecabStatus(), {
    available: true,
    enabled: true,
    path: "/usr/bin/mecab",
  });
  assert.equal(deps.getCurrentSecondarySub(), "secondary");
  assert.deepEqual(await deps.tokenizeCurrentSubtitle(), { text: "x" });
});

test("createIpcDepsRuntimeService handles missing optional runtime resources", () => {
  const deps = createIpcDepsRuntimeService({
    getInvisibleWindow: () => null,
    getMainWindow: () => null,
    getVisibleOverlayVisibility: () => false,
    getInvisibleOverlayVisibility: () => false,
    onOverlayModalClosed: () => {},
    openYomitanSettings: () => {},
    quitApp: () => {},
    toggleVisibleOverlay: () => {},
    tokenizeCurrentSubtitle: async () => null,
    getCurrentSubtitleAss: () => "",
    getMpvSubtitleRenderMetrics: () => null,
    getSubtitlePosition: () => null,
    getSubtitleStyle: () => null,
    saveSubtitlePosition: () => {},
    getMecabTokenizer: () => null,
    handleMpvCommand: () => {},
    getKeybindings: () => null,
    getSecondarySubMode: () => "hidden",
    getMpvClient: () => null,
    runSubsyncManual: async () => ({ ok: false }),
    getAnkiConnectStatus: () => false,
    getRuntimeOptions: () => null,
    setRuntimeOption: () => ({ ok: false }),
    cycleRuntimeOption: () => ({ ok: false }),
  });

  deps.setInvisibleIgnoreMouseEvents(true, { forward: true });
  deps.toggleDevTools();
  deps.setMecabEnabled(true);

  assert.deepEqual(deps.getMecabStatus(), {
    available: false,
    enabled: false,
    path: null,
  });
  assert.equal(deps.getCurrentSecondarySub(), "");
});

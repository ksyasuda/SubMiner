import test from "node:test";
import assert from "node:assert/strict";
import {
  createInitializeOverlayRuntimeDepsService,
  createInvisibleOverlayVisibilityDepsRuntimeService,
  createOverlayWindowRuntimeDepsService,
  createVisibleOverlayVisibilityDepsRuntimeService,
} from "./overlay-runtime-deps-service";

test("createOverlayWindowRuntimeDepsService maps runtime state providers", () => {
  let visible = true;
  let invisible = false;
  const deps = createOverlayWindowRuntimeDepsService({
    isDev: false,
    getOverlayDebugVisualizationEnabled: () => true,
    ensureOverlayWindowLevel: () => {},
    onRuntimeOptionsChanged: () => {},
    setOverlayDebugVisualizationEnabled: () => {},
    getVisibleOverlayVisible: () => visible,
    getInvisibleOverlayVisible: () => invisible,
    tryHandleOverlayShortcutLocalFallback: () => false,
    onWindowClosed: () => {},
  });

  assert.equal(deps.isOverlayVisible("visible"), true);
  assert.equal(deps.isOverlayVisible("invisible"), false);
  visible = false;
  invisible = true;
  assert.equal(deps.isOverlayVisible("visible"), false);
  assert.equal(deps.isOverlayVisible("invisible"), true);
});

test("createInitializeOverlayRuntimeDepsService passes through overlay init deps", () => {
  const windows: any[] = [];
  const deps = createInitializeOverlayRuntimeDepsService({
    backendOverride: null,
    getInitialInvisibleOverlayVisibility: () => true,
    createMainWindow: () => {},
    createInvisibleWindow: () => {},
    registerGlobalShortcuts: () => {},
    updateOverlayBounds: () => {},
    isVisibleOverlayVisible: () => false,
    isInvisibleOverlayVisible: () => true,
    updateVisibleOverlayVisibility: () => {},
    updateInvisibleOverlayVisibility: () => {},
    getOverlayWindows: () => windows as never,
    syncOverlayShortcuts: () => {},
    setWindowTracker: () => {},
    getResolvedConfig: () => ({ ankiConnect: undefined }),
    getSubtitleTimingTracker: () => null,
    getMpvClient: () => null,
    getRuntimeOptionsManager: () => null,
    setAnkiIntegration: () => {},
    showDesktopNotification: () => {},
    createFieldGroupingCallback: () => async () => ({
      keepNoteId: 0,
      deleteNoteId: 0,
      deleteDuplicate: true,
      cancelled: true,
    }),
  });

  assert.equal(deps.getInitialInvisibleOverlayVisibility(), true);
  assert.equal(deps.getOverlayWindows().length, 0);
});

test("createVisibleOverlayVisibilityDepsRuntimeService snapshots runtime values", () => {
  const deps = createVisibleOverlayVisibilityDepsRuntimeService({
    getVisibleOverlayVisible: () => true,
    getMainWindow: () => null,
    getWindowTracker: () => null,
    getTrackerNotReadyWarningShown: () => false,
    setTrackerNotReadyWarningShown: () => {},
    shouldBindVisibleOverlayToMpvSubVisibility: () => true,
    getPreviousSecondarySubVisibility: () => null,
    setPreviousSecondarySubVisibility: () => {},
    isMpvConnected: () => false,
    mpvSend: () => {},
    secondarySubVisibilityRequestId: 123,
    updateOverlayBounds: () => {},
    ensureOverlayWindowLevel: () => {},
    enforceOverlayLayerOrder: () => {},
    syncOverlayShortcuts: () => {},
  });

  assert.equal(deps.visibleOverlayVisible, true);
  assert.equal(deps.secondarySubVisibilityRequestId, 123);
  assert.equal(deps.mpvConnected, false);
});

test("createInvisibleOverlayVisibilityDepsRuntimeService snapshots runtime values", () => {
  const deps = createInvisibleOverlayVisibilityDepsRuntimeService({
    getInvisibleWindow: () => null,
    getVisibleOverlayVisible: () => true,
    getInvisibleOverlayVisible: () => false,
    getWindowTracker: () => null,
    updateOverlayBounds: () => {},
    ensureOverlayWindowLevel: () => {},
    enforceOverlayLayerOrder: () => {},
    syncOverlayShortcuts: () => {},
  });

  assert.equal(deps.visibleOverlayVisible, true);
  assert.equal(deps.invisibleOverlayVisible, false);
});

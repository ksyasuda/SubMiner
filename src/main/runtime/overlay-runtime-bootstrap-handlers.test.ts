import assert from 'node:assert/strict';
import test from 'node:test';
import type { BaseWindowTracker } from '../../window-trackers';
import type { KikuFieldGroupingChoice } from '../../types';
import { createOverlayRuntimeBootstrapHandlers } from './overlay-runtime-bootstrap-handlers';

test('overlay runtime bootstrap handlers compose options builder and bootstrap handler', () => {
  const appState = {
    backendOverride: null as string | null,
    windowTracker: null as BaseWindowTracker | null,
    subtitleTimingTracker: null as unknown,
    mpvClient: null,
    mpvSocketPath: '/tmp/mpv.sock',
    runtimeOptionsManager: null,
    ankiIntegration: null as unknown,
  };
  let initialized = false;
  let invisibleOverlayVisible = false;
  let warmupsStarted = 0;

  const { initializeOverlayRuntime } = createOverlayRuntimeBootstrapHandlers({
    initializeOverlayRuntimeMainDeps: {
      appState,
      overlayManager: {
        getVisibleOverlayVisible: () => true,
        getInvisibleOverlayVisible: () => false,
      },
      overlayVisibilityRuntime: {
        updateVisibleOverlayVisibility: () => {},
        updateInvisibleOverlayVisibility: () => {},
      },
      overlayShortcutsRuntime: {
        syncOverlayShortcuts: () => {},
      },
      getInitialInvisibleOverlayVisibility: () => false,
      createMainWindow: () => {},
      createInvisibleWindow: () => {},
      registerGlobalShortcuts: () => {},
      updateVisibleOverlayBounds: () => {},
      updateInvisibleOverlayBounds: () => {},
      getOverlayWindows: () => [],
      getResolvedConfig: () => ({}),
      showDesktopNotification: () => {},
      createFieldGroupingCallback: () => async () =>
        ({
          keepNoteId: 1,
          deleteNoteId: 2,
          deleteDuplicate: false,
          cancelled: true,
        }) as KikuFieldGroupingChoice,
      getKnownWordCacheStatePath: () => '/tmp/known.json',
    },
    initializeOverlayRuntimeBootstrapDeps: {
      isOverlayRuntimeInitialized: () => initialized,
      initializeOverlayRuntimeCore: () => ({ invisibleOverlayVisible: true }),
      setInvisibleOverlayVisible: (visible) => {
        invisibleOverlayVisible = visible;
      },
      setOverlayRuntimeInitialized: (next) => {
        initialized = next;
      },
      startBackgroundWarmups: () => {
        warmupsStarted += 1;
      },
    },
  });

  initializeOverlayRuntime();
  initializeOverlayRuntime();

  assert.equal(invisibleOverlayVisible, true);
  assert.equal(initialized, true);
  assert.equal(warmupsStarted, 1);
});

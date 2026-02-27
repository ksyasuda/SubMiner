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
  let warmupsStarted = 0;

  const { initializeOverlayRuntime } = createOverlayRuntimeBootstrapHandlers({
    initializeOverlayRuntimeMainDeps: {
      appState,
      overlayManager: {
        getVisibleOverlayVisible: () => true,
      },
      overlayVisibilityRuntime: {
        updateVisibleOverlayVisibility: () => {},
      },
      overlayShortcutsRuntime: {
        syncOverlayShortcuts: () => {},
      },
      createMainWindow: () => {},
      registerGlobalShortcuts: () => {},
      updateVisibleOverlayBounds: () => {},
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
      initializeOverlayRuntimeCore: () => {},
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

  assert.equal(initialized, true);
  assert.equal(warmupsStarted, 1);
});

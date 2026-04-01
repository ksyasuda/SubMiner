import assert from 'node:assert/strict';
import test from 'node:test';
import type { OverlayHostedModal } from '../shared/ipc/contracts';

import { createOverlayUiBootstrapRuntimeInput } from './overlay-ui-bootstrap-runtime-input';

test('overlay ui bootstrap runtime input builder preserves grouped wiring', () => {
  const input = createOverlayUiBootstrapRuntimeInput({
    windows: {
      state: {
        getMainWindow: () => null,
        setMainWindow: () => {},
        getModalWindow: () => null,
        setModalWindow: () => {},
        getVisibleOverlayVisible: () => false,
        setVisibleOverlayVisible: () => {},
        getOverlayDebugVisualizationEnabled: () => false,
        setOverlayDebugVisualizationEnabled: () => {},
      },
      geometry: {
        getCurrentOverlayGeometry: () => ({ x: 1, y: 2, width: 3, height: 4 }),
      },
      modal: {
        setModalWindowBounds: () => {},
        onModalStateChange: () => {},
      },
      modalRuntime: {
        handleOverlayModalClosed: (_modal: OverlayHostedModal) => {},
        notifyOverlayModalOpened: (_modal: OverlayHostedModal) => {},
        waitForModalOpen: async () => true,
        getRestoreVisibleOverlayOnModalClose: () => new Set<OverlayHostedModal>(),
        openRuntimeOptionsPalette: () => {},
        sendToActiveOverlayWindow: () => false,
      },
      visibility: {
        service: {
          getModalActive: () => false,
          getForceMousePassthrough: () => false,
          getWindowTracker: () => null,
          getTrackerNotReadyWarningShown: () => false,
          setTrackerNotReadyWarningShown: () => {},
          updateVisibleOverlayBounds: () => {},
          ensureOverlayWindowLevel: () => {},
          syncPrimaryOverlayWindowLayer: () => {},
          enforceOverlayLayerOrder: () => {},
          syncOverlayShortcuts: () => {},
          isMacOSPlatform: () => false,
          isWindowsPlatform: () => false,
          showOverlayLoadingOsd: () => {},
          resolveFallbackBounds: () => ({ x: 5, y: 6, width: 7, height: 8 }),
        },
        overlayWindows: {
          createOverlayWindowCore: () => ({ isDestroyed: () => false }),
          isDev: false,
          ensureOverlayWindowLevel: () => {},
          onRuntimeOptionsChanged: () => {},
          setOverlayDebugVisualizationEnabled: () => {},
          isOverlayVisible: () => false,
          getYomitanSession: () => null,
          tryHandleOverlayShortcutLocalFallback: () => false,
          forwardTabToMpv: () => {},
          onWindowClosed: () => {},
        },
        actions: {
          setVisibleOverlayVisibleCore: () => {},
        },
      },
    },
    overlayActions: {
      getRuntimeOptionsManager: () => null,
      getMpvClient: () => null,
      broadcastRuntimeOptionsChangedRuntime: () => {},
      broadcastToOverlayWindows: () => {},
      setOverlayDebugVisualizationEnabledRuntime: () => {},
    },
    tray: null,
    bootstrap: {
      initializeOverlayRuntimeMainDeps: {
        appState: {
          backendOverride: null,
          windowTracker: null,
          subtitleTimingTracker: null,
          mpvClient: null,
          mpvSocketPath: '/tmp/mpv.sock',
          runtimeOptionsManager: null,
          ankiIntegration: null,
        },
        overlayManager: {
          getVisibleOverlayVisible: () => false,
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
        getResolvedConfig: () => ({}) as never,
        showDesktopNotification: () => {},
        createFieldGroupingCallback: () => () => Promise.resolve({} as never),
        getKnownWordCacheStatePath: () => '/tmp/known.json',
        shouldStartAnkiIntegration: () => false,
      },
      initializeOverlayRuntimeBootstrapDeps: {
        isOverlayRuntimeInitialized: () => false,
        initializeOverlayRuntimeCore: () => {},
        setOverlayRuntimeInitialized: () => {},
        startBackgroundWarmups: () => {},
      },
      onInitialized: () => {},
    },
    runtimeState: {
      isOverlayRuntimeInitialized: () => false,
      setOverlayRuntimeInitialized: () => {},
    },
    mpvSubtitle: {
      ensureOverlayMpvSubtitlesHidden: () => {},
      syncOverlayMpvSubtitleSuppression: () => {},
    },
  });

  assert.equal(input.tray, null);
  assert.equal(input.windows.windowState.getMainWindow(), null);
  assert.equal(input.windows.geometry.getCurrentOverlayGeometry().width, 3);
  assert.equal(input.windows.visibilityService.resolveFallbackBounds().height, 8);
  assert.equal(
    input.bootstrap.initializeOverlayRuntimeBootstrapDeps.isOverlayRuntimeInitialized(),
    false,
  );
});

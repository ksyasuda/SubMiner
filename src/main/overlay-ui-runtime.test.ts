import assert from 'node:assert/strict';
import test from 'node:test';
import type { OverlayHostedModal } from '../shared/ipc/contracts';

import { createOverlayUiRuntime } from './overlay-ui-runtime';

type MockWindow = {
  destroyed: boolean;
  isDestroyed: () => boolean;
};

function createWindow(): MockWindow {
  return {
    destroyed: false,
    isDestroyed() {
      return this.destroyed;
    },
  };
}

test('overlay ui runtime lazy-creates main window for toggle visibility actions', async () => {
  const calls: string[] = [];
  let mainWindow: MockWindow | null = null;
  const createdWindow = createWindow();
  let visibleOverlayVisible = false;

  const overlayUi = createOverlayUiRuntime({
    windows: {
      windowState: {
        getMainWindow: () => mainWindow,
        setMainWindow: (window) => {
          mainWindow = window;
        },
        getModalWindow: () => null,
        setModalWindow: () => {},
        getVisibleOverlayVisible: () => visibleOverlayVisible,
        setVisibleOverlayVisible: (visible) => {
          visibleOverlayVisible = visible;
        },
        getOverlayDebugVisualizationEnabled: () => false,
        setOverlayDebugVisualizationEnabled: () => {},
      },
      geometry: {
        getCurrentOverlayGeometry: () => ({ x: 0, y: 0, width: 100, height: 100 }),
      },
      modal: {
        onModalStateChange: () => {},
      },
      modalRuntime: {
        handleOverlayModalClosed: () => {},
        notifyOverlayModalOpened: () => {},
        waitForModalOpen: async () => false,
        getRestoreVisibleOverlayOnModalClose: () => new Set<OverlayHostedModal>(),
        openRuntimeOptionsPalette: () => {},
        sendToActiveOverlayWindow: () => false,
      },
      visibilityService: {
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
        resolveFallbackBounds: () => ({ x: 0, y: 0, width: 100, height: 100 }),
      },
      overlayWindows: {
        createOverlayWindowCore: () => createdWindow,
        isDev: false,
        ensureOverlayWindowLevel: () => {},
        onRuntimeOptionsChanged: () => {},
        setOverlayDebugVisualizationEnabled: () => {},
        isOverlayVisible: () => visibleOverlayVisible,
        getYomitanSession: () => null,
        tryHandleOverlayShortcutLocalFallback: () => false,
        forwardTabToMpv: () => {},
        onWindowClosed: () => {},
      },
      visibilityActions: {
        setVisibleOverlayVisibleCore: ({ visible, setVisibleOverlayVisibleState }) => {
          calls.push(`setVisible:${visible}`);
          setVisibleOverlayVisibleState(visible);
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
          getVisibleOverlayVisible: () => visibleOverlayVisible,
        },
        overlayVisibilityRuntime: {
          updateVisibleOverlayVisibility: () => {},
        },
        overlayShortcutsRuntime: {
          syncOverlayShortcuts: () => {},
        },
        createMainWindow: () => {
          calls.push('bootstrapCreateMainWindow');
        },
        registerGlobalShortcuts: () => {},
        updateVisibleOverlayBounds: () => {},
        getOverlayWindows: () => [],
        getResolvedConfig: () => ({ ankiConnect: {} }) as never,
        showDesktopNotification: () => {},
        createFieldGroupingCallback: () => () => Promise.resolve({} as never),
        getKnownWordCacheStatePath: () => '/tmp/known.json',
        shouldStartAnkiIntegration: () => false,
      },
      initializeOverlayRuntimeBootstrapDeps: {
        isOverlayRuntimeInitialized: () => true,
        initializeOverlayRuntimeCore: () => {},
        setOverlayRuntimeInitialized: () => {},
        startBackgroundWarmups: () => {},
      },
      onInitialized: () => {},
    },
    runtimeState: {
      isOverlayRuntimeInitialized: () => true,
      setOverlayRuntimeInitialized: () => {},
    },
    mpvSubtitle: {
      ensureOverlayMpvSubtitlesHidden: async () => {
        calls.push('hideMpvSubs');
      },
      syncOverlayMpvSubtitleSuppression: () => {
        calls.push('syncMpvSubs');
      },
    },
  });

  overlayUi.toggleVisibleOverlay();

  assert.equal(mainWindow, createdWindow);
  assert.deepEqual(calls, ['hideMpvSubs', 'setVisible:true', 'syncMpvSubs']);
});

test('overlay ui runtime initializes overlay runtime before visible action when needed', async () => {
  const calls: string[] = [];
  let visibleOverlayVisible = false;
  let overlayRuntimeInitialized = false;

  const overlayUi = createOverlayUiRuntime({
    windowState: {
      getMainWindow: () => null,
      setMainWindow: () => {},
      getModalWindow: () => null,
      setModalWindow: () => {},
      getVisibleOverlayVisible: () => visibleOverlayVisible,
      setVisibleOverlayVisible: (visible) => {
        visibleOverlayVisible = visible;
      },
      getOverlayDebugVisualizationEnabled: () => false,
      setOverlayDebugVisualizationEnabled: () => {},
    },
    geometry: {
      getCurrentOverlayGeometry: () => ({ x: 0, y: 0, width: 100, height: 100 }),
    },
    modal: {
      onModalStateChange: () => {},
    },
    modalRuntime: {
      handleOverlayModalClosed: () => {},
      notifyOverlayModalOpened: () => {},
      waitForModalOpen: async () => false,
      getRestoreVisibleOverlayOnModalClose: () => new Set<OverlayHostedModal>(),
      openRuntimeOptionsPalette: () => {},
      sendToActiveOverlayWindow: () => false,
    },
    visibilityService: {
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
      resolveFallbackBounds: () => ({ x: 0, y: 0, width: 100, height: 100 }),
    },
    overlayWindows: {
      createOverlayWindowCore: () => createWindow(),
      isDev: false,
      ensureOverlayWindowLevel: () => {},
      onRuntimeOptionsChanged: () => {},
      setOverlayDebugVisualizationEnabled: () => {},
      isOverlayVisible: () => visibleOverlayVisible,
      getYomitanSession: () => null,
      tryHandleOverlayShortcutLocalFallback: () => false,
      forwardTabToMpv: () => {},
      onWindowClosed: () => {},
    },
    visibilityActions: {
      setVisibleOverlayVisibleCore: ({ visible, setVisibleOverlayVisibleState }) => {
        calls.push(`setVisible:${visible}`);
        setVisibleOverlayVisibleState(visible);
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
          getVisibleOverlayVisible: () => visibleOverlayVisible,
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
        getResolvedConfig: () => ({ ankiConnect: {} }) as never,
        showDesktopNotification: () => {},
        createFieldGroupingCallback: () => () => Promise.resolve({} as never),
        getKnownWordCacheStatePath: () => '/tmp/known.json',
        shouldStartAnkiIntegration: () => false,
      },
      initializeOverlayRuntimeBootstrapDeps: {
        isOverlayRuntimeInitialized: () => overlayRuntimeInitialized,
        initializeOverlayRuntimeCore: () => {
          calls.push('initializeOverlayRuntimeCore');
        },
        setOverlayRuntimeInitialized: (initialized) => {
          overlayRuntimeInitialized = initialized;
          calls.push(`setInitialized:${initialized}`);
        },
        startBackgroundWarmups: () => {
          calls.push('startBackgroundWarmups');
        },
      },
      onInitialized: () => {
        calls.push('onInitialized');
      },
    },
    runtimeState: {
      isOverlayRuntimeInitialized: () => overlayRuntimeInitialized,
      setOverlayRuntimeInitialized: (initialized) => {
        overlayRuntimeInitialized = initialized;
      },
    },
    mpvSubtitle: {
      ensureOverlayMpvSubtitlesHidden: async () => {
        calls.push('hideMpvSubs');
      },
      syncOverlayMpvSubtitleSuppression: () => {
        calls.push('syncMpvSubs');
      },
    },
  });

  overlayUi.setVisibleOverlayVisible(true);

  assert.deepEqual(calls, [
    'setInitialized:true',
    'initializeOverlayRuntimeCore',
    'startBackgroundWarmups',
    'onInitialized',
    'syncMpvSubs',
    'hideMpvSubs',
    'setVisible:true',
    'syncMpvSubs',
  ]);
});

test('overlay ui runtime delegates modal actions to injected modal runtime', async () => {
  const calls: string[] = [];
  const restoreOnClose = new Set<OverlayHostedModal>();

  const overlayUi = createOverlayUiRuntime({
    windowState: {
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
      getCurrentOverlayGeometry: () => ({ x: 0, y: 0, width: 100, height: 100 }),
    },
    modal: {
      onModalStateChange: () => {},
    },
    visibilityService: {
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
      resolveFallbackBounds: () => ({ x: 0, y: 0, width: 100, height: 100 }),
    },
    overlayWindows: {
      createOverlayWindowCore: () => createWindow(),
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
    visibilityActions: {
      setVisibleOverlayVisibleCore: ({ visible, setVisibleOverlayVisibleState }) => {
        setVisibleOverlayVisibleState(visible);
      },
    },
    overlayActions: {
      getRuntimeOptionsManager: () => null,
      getMpvClient: () => null,
      broadcastRuntimeOptionsChangedRuntime: () => {},
      broadcastToOverlayWindows: () => {},
      setOverlayDebugVisualizationEnabledRuntime: () => {},
    },
    modalRuntime: {
      sendToActiveOverlayWindow: (channel, payload, runtimeOptions) => {
        calls.push(`send:${channel}:${String(payload)}`);
        if (runtimeOptions?.restoreOnModalClose) {
          restoreOnClose.add(runtimeOptions.restoreOnModalClose);
        }
        return true;
      },
      openRuntimeOptionsPalette: () => {
        calls.push('openRuntimeOptionsPalette');
      },
      handleOverlayModalClosed: (modal) => {
        calls.push(`closed:${modal}`);
      },
      notifyOverlayModalOpened: (modal) => {
        calls.push(`opened:${modal}`);
      },
      waitForModalOpen: async (modal, timeoutMs) => {
        calls.push(`wait:${modal}:${timeoutMs}`);
        return true;
      },
      getRestoreVisibleOverlayOnModalClose: () => restoreOnClose,
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
        getResolvedConfig: () => ({ ankiConnect: {} }) as never,
        showDesktopNotification: () => {},
        createFieldGroupingCallback: () => () => Promise.resolve({} as never),
        getKnownWordCacheStatePath: () => '/tmp/known.json',
        shouldStartAnkiIntegration: () => false,
      },
      initializeOverlayRuntimeBootstrapDeps: {
        isOverlayRuntimeInitialized: () => true,
        initializeOverlayRuntimeCore: () => {},
        setOverlayRuntimeInitialized: () => {},
        startBackgroundWarmups: () => {},
      },
    },
    runtimeState: {
      isOverlayRuntimeInitialized: () => true,
      setOverlayRuntimeInitialized: () => {},
    },
    mpvSubtitle: {
      ensureOverlayMpvSubtitlesHidden: async () => {},
      syncOverlayMpvSubtitleSuppression: () => {},
    },
  });

  assert.equal(
    overlayUi.sendToActiveOverlayWindow('jimaku:open', 'payload', {
      restoreOnModalClose: 'jimaku',
    }),
    true,
  );
  overlayUi.openRuntimeOptionsPalette();
  overlayUi.notifyOverlayModalOpened('runtime-options');
  overlayUi.handleOverlayModalClosed('runtime-options');
  assert.equal(await overlayUi.waitForModalOpen('youtube-track-picker', 50), true);
  assert.equal(overlayUi.getRestoreVisibleOverlayOnModalClose(), restoreOnClose);
  assert.deepEqual(calls, [
    'send:jimaku:open:payload',
    'openRuntimeOptionsPalette',
    'opened:runtime-options',
    'closed:runtime-options',
    'wait:youtube-track-picker:50',
  ]);
});

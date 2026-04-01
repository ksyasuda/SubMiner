import assert from 'node:assert/strict';
import test from 'node:test';

import { createIpcRuntime, createIpcRuntimeFromMainState } from './ipc-runtime';

function createBaseRuntimeInput(capturedRegistration: { value: unknown | null }) {
  const manualResult = { ok: true, summary: 'done' };
  const main = {
    window: {
      getMainWindow: () => null,
      getVisibleOverlayVisibility: () => false,
      focusMainWindow: () => {},
      onOverlayModalClosed: () => {},
      onOverlayModalOpened: () => {},
      onYoutubePickerResolve: async () => ({ ok: true }) as never,
      openYomitanSettings: () => {},
      quitApp: () => {},
      toggleVisibleOverlay: () => {},
    },
    subtitle: {
      tokenizeCurrentSubtitle: async () => null,
      getCurrentSubtitleRaw: () => '',
      getCurrentSubtitleAss: () => '',
      getSubtitleSidebarSnapshot: async () => null as never,
      getPlaybackPaused: () => false,
      getSubtitlePosition: () => null,
      getSubtitleStyle: () => null,
      saveSubtitlePosition: () => {},
      getMecabTokenizer: () => null,
      getKeybindings: () => [],
      getConfiguredShortcuts: () => null,
      getStatsToggleKey: () => '',
      getMarkWatchedKey: () => '',
      getSecondarySubMode: () => 'hover',
    },
    controller: {
      getControllerConfig: () => ({}) as never,
      saveControllerConfig: () => {},
      saveControllerPreference: () => {},
    },
    runtime: {
      getMpvClient: () => null,
      getAnkiConnectStatus: () => false,
      getRuntimeOptions: () => [],
      reportOverlayContentBounds: () => {},
      getImmersionTracker: () => null,
    },
    anilist: {
      getStatus: () => null,
      clearToken: () => {},
      openSetup: () => {},
      getQueueStatus: () => null,
      retryQueueNow: async () => ({ ok: true, message: 'ok' }) as never,
    },
    mining: {
      appendClipboardVideoToQueue: () => ({ ok: true, message: 'ok' }),
    },
  };
  const ankiJimaku = {
    patchAnkiConnectEnabled: () => {},
    getResolvedConfig: () => ({}),
    getRuntimeOptionsManager: () => null,
    getSubtitleTimingTracker: () => null,
    getMpvClient: () => null,
    getAnkiIntegration: () => null,
    setAnkiIntegration: () => {},
    getKnownWordCacheStatePath: () => '/tmp/known-words.json',
    showDesktopNotification: () => {},
    createFieldGroupingCallback: () => async () => ({}) as never,
    broadcastRuntimeOptionsChanged: () => {},
    getFieldGroupingResolver: () => null,
    setFieldGroupingResolver: () => {},
    parseMediaInfo: () => ({}) as never,
    getCurrentMediaPath: () => null,
    jimakuFetchJson: async () => ({ data: null, error: null }) as never,
    getJimakuMaxEntryResults: () => 5,
    getJimakuLanguagePreference: () => 'ja' as const,
    resolveJimakuApiKey: async () => null,
    isRemoteMediaPath: () => false,
    downloadToFile: async () => ({ ok: true, path: '/tmp/file' }) as never,
  };

  return {
    mpv: {
      mainDeps: {
        triggerSubsyncFromConfig: () => {},
        openRuntimeOptionsPalette: () => {},
        openYoutubeTrackPicker: () => {},
        cycleRuntimeOption: () => ({ ok: true }),
        showMpvOsd: () => {},
        replayCurrentSubtitle: () => {},
        playNextSubtitle: () => {},
        shiftSubDelayToAdjacentSubtitle: async () => {},
        sendMpvCommand: () => {},
        getMpvClient: () => null,
        isMpvConnected: () => true,
        hasRuntimeOptionsManager: () => true,
      },
      handleMpvCommandFromIpcRuntime: () => {},
      runSubsyncManualFromIpc: async () => manualResult as never,
    },
    main,
    ankiJimaku,
    registration: {
      runtimeOptions: {
        getRuntimeOptionsManager: () => null,
        showMpvOsd: () => {},
      },
      main,
      ankiJimaku,
      registerIpcRuntimeServices: (params: unknown) => {
        capturedRegistration.value = params;
      },
    },
    registerIpcRuntimeServices: (params: unknown) => {
      capturedRegistration.value = params;
    },
    manualResult,
  };
}

test('ipc runtime registers composed IPC handlers from explicit registration input', async () => {
  const capturedRegistration = { value: null as unknown | null };
  const input = createBaseRuntimeInput(capturedRegistration);

  const runtime = createIpcRuntime({
    mpv: input.mpv,
    registration: input.registration,
  });

  runtime.registerIpcRuntimeHandlers();

  assert.ok(capturedRegistration.value);
  const registration = capturedRegistration.value as {
    runtimeOptions: { showMpvOsd: unknown };
    mainDeps: {
      handleMpvCommand: unknown;
      runSubsyncManual: (payload: unknown) => Promise<unknown>;
    };
  };
  assert.equal(registration.runtimeOptions.showMpvOsd !== undefined, true);
  assert.equal(registration.mainDeps.handleMpvCommand instanceof Function, true);
  assert.deepEqual(
    await registration.mainDeps.runSubsyncManual({ payload: null } as never),
    input.manualResult,
  );
});

test('ipc runtime builds grouped registration input from main state', async () => {
  const capturedRegistration = { value: null as unknown | null };
  const input = createBaseRuntimeInput(capturedRegistration);

  const runtime = createIpcRuntimeFromMainState({
    mpv: input.mpv,
    runtimeOptions: input.registration.runtimeOptions,
    main: input.main,
    ankiJimaku: input.ankiJimaku,
  });

  runtime.registerIpcRuntimeHandlers();

  assert.ok(capturedRegistration.value);
  const registration = capturedRegistration.value as {
    runtimeOptions: { showMpvOsd: unknown };
    mainDeps: {
      handleMpvCommand: unknown;
      runSubsyncManual: (payload: unknown) => Promise<unknown>;
    };
  };
  assert.equal(registration.runtimeOptions.showMpvOsd !== undefined, true);
  assert.equal(registration.mainDeps.handleMpvCommand instanceof Function, true);
  assert.deepEqual(
    await registration.mainDeps.runSubsyncManual({ payload: null } as never),
    input.manualResult,
  );
});

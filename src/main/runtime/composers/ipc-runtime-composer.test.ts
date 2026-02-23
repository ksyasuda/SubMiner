import assert from 'node:assert/strict';
import test from 'node:test';
import { composeIpcRuntimeHandlers } from './ipc-runtime-composer';

test('composeIpcRuntimeHandlers returns callable IPC handlers and registration bridge', async () => {
  let registered = false;
  let receivedSourceTrackId: number | null | undefined;

  const composed = composeIpcRuntimeHandlers({
    mpvCommandMainDeps: {
      triggerSubsyncFromConfig: async () => {},
      openRuntimeOptionsPalette: () => {},
      cycleRuntimeOption: () => ({ ok: true }),
      showMpvOsd: () => {},
      replayCurrentSubtitle: () => {},
      playNextSubtitle: () => {},
      sendMpvCommand: () => {},
      isMpvConnected: () => false,
      hasRuntimeOptionsManager: () => true,
    },
    handleMpvCommandFromIpcRuntime: () => {},
    runSubsyncManualFromIpc: async (request) => {
      receivedSourceTrackId = request.sourceTrackId;
      return {
        ok: true,
        message: 'ok',
      };
    },
    registration: {
      runtimeOptions: {
        getRuntimeOptionsManager: () => null,
        showMpvOsd: () => {},
      },
      mainDeps: {
        getInvisibleWindow: () => null,
        getMainWindow: () => null,
        getVisibleOverlayVisibility: () => false,
        getInvisibleOverlayVisibility: () => false,
        focusMainWindow: () => {},
        onOverlayModalClosed: () => {},
        openYomitanSettings: () => {},
        quitApp: () => {},
        toggleVisibleOverlay: () => {},
        tokenizeCurrentSubtitle: async () => null,
        getCurrentSubtitleRaw: () => '',
        getCurrentSubtitleAss: () => '',
        getMpvSubtitleRenderMetrics: () => ({}) as never,
        getSubtitlePosition: () => ({}) as never,
        getSubtitleStyle: () => ({}) as never,
        saveSubtitlePosition: () => {},
        getMecabTokenizer: () => null,
        getKeybindings: () => [],
        getConfiguredShortcuts: () => ({}) as never,
        getSecondarySubMode: () => 'hover' as never,
        getMpvClient: () => null,
        getAnkiConnectStatus: () => false,
        getRuntimeOptions: () => [],
        reportOverlayContentBounds: () => {},
        reportHoveredSubtitleToken: () => {},
        getAnilistStatus: () => ({}) as never,
        clearAnilistToken: () => {},
        openAnilistSetup: () => {},
        getAnilistQueueStatus: () => ({}) as never,
        retryAnilistQueueNow: async () => ({ ok: true, message: 'ok' }),
        appendClipboardVideoToQueue: () => ({ ok: true, message: 'ok' }),
      },
      ankiJimakuDeps: {
        patchAnkiConnectEnabled: () => {},
        getResolvedConfig: () => ({}) as never,
        getRuntimeOptionsManager: () => null,
        getSubtitleTimingTracker: () => null,
        getMpvClient: () => null,
        getAnkiIntegration: () => null,
        setAnkiIntegration: () => {},
        getKnownWordCacheStatePath: () => '',
        showDesktopNotification: () => {},
        createFieldGroupingCallback: () => (() => {}) as never,
        broadcastRuntimeOptionsChanged: () => {},
        getFieldGroupingResolver: () => null,
        setFieldGroupingResolver: () => {},
        parseMediaInfo: () => ({}) as never,
        getCurrentMediaPath: () => null,
        jimakuFetchJson: async () => ({ data: null }) as never,
        getJimakuMaxEntryResults: () => 0,
        getJimakuLanguagePreference: () => 'ja' as never,
        resolveJimakuApiKey: async () => null,
        isRemoteMediaPath: () => false,
        downloadToFile: async () => ({ ok: true, path: '/tmp/file' }),
      },
      registerIpcRuntimeServices: () => {
        registered = true;
      },
    },
  });

  assert.equal(typeof composed.handleMpvCommandFromIpc, 'function');
  assert.equal(typeof composed.runSubsyncManualFromIpc, 'function');
  assert.equal(typeof composed.registerIpcRuntimeHandlers, 'function');

  const result = await composed.runSubsyncManualFromIpc({
    engine: 'alass',
    sourceTrackId: 7,
  });
  assert.deepEqual(result, { ok: true, message: 'ok' });
  assert.equal(receivedSourceTrackId, 7);

  composed.registerIpcRuntimeHandlers();
  assert.equal(registered, true);
});

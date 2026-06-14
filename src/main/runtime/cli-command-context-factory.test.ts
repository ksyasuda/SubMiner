import assert from 'node:assert/strict';
import test from 'node:test';
import { createCliCommandContextFactory } from './cli-command-context-factory';

test('cli command context factory composes main deps and context handlers', () => {
  const calls: string[] = [];
  const appState = {
    mpvSocketPath: '/tmp/mpv.sock',
    mpvClient: null,
    texthookerPort: 5174,
    overlayRuntimeInitialized: false,
  };

  const createContext = createCliCommandContextFactory({
    appState,
    texthookerService: { isRunning: () => false, start: () => null },
    getResolvedConfig: () => ({
      texthooker: { openBrowser: true },
      annotationWebsocket: { enabled: true, port: 6678 },
    }),
    defaultWebsocketPort: 6677,
    defaultAnnotationWebsocketPort: 6678,
    hasMpvWebsocketPlugin: () => false,
    openExternal: async () => {},
    logBrowserOpenError: () => {},
    showMpvOsd: (text) => calls.push(`osd:${text}`),
    initializeOverlayRuntime: () => calls.push('init-overlay'),
    toggleVisibleOverlay: () => calls.push('toggle-visible'),
    togglePrimarySubtitleBar: () => calls.push('toggle-primary-subtitle'),
    openFirstRunSetupWindow: () => calls.push('setup'),
    setVisibleOverlayVisible: (visible) => calls.push(`set-visible:${visible}`),
    copyCurrentSubtitle: () => calls.push('copy-sub'),
    startPendingMultiCopy: (timeoutMs) => calls.push(`multi:${timeoutMs}`),
    mineSentenceCard: async () => {},
    startPendingMineSentenceMultiple: () => {},
    updateLastCardFromClipboard: async () => {},
    refreshKnownWordCache: async () => {},
    triggerFieldGrouping: async () => {},
    triggerSubsyncFromConfig: async () => {},
    markLastCardAsAudioCard: async () => {},
    dispatchSessionAction: async () => {},
    getAnilistStatus: () => ({
      tokenStatus: 'resolved',
      tokenSource: 'literal',
      tokenMessage: null,
      tokenResolvedAt: null,
      tokenErrorAt: null,
      queuePending: 0,
      queueReady: 0,
      queueDeadLetter: 0,
      queueLastAttemptAt: null,
      queueLastError: null,
    }),
    clearAnilistToken: () => {},
    openAnilistSetupWindow: () => {},
    openJellyfinSetupWindow: () => {},
    getAnilistQueueStatus: () => ({
      pending: 0,
      ready: 0,
      deadLetter: 0,
      lastAttemptAt: null,
      lastError: null,
    }),
    processNextAnilistRetryUpdate: async () => ({ ok: true, message: 'ok' }),
    generateCharacterDictionary: async () => ({
      zipPath: '/tmp/anilist-1.zip',
      fromCache: false,
      mediaId: 1,
      mediaTitle: 'Test',
      entryCount: 10,
    }),
    runStatsCommand: async () => {},
    runJellyfinCommand: async () => {},
    runUpdateCommand: async () => {},
    runEnsureLinuxRuntimePluginAssetsCommand: async () => {},
    runYoutubePlaybackFlow: async () => {},
    openYomitanSettings: () => {},
    openConfigSettingsWindow: () => {},
    cycleSecondarySubMode: () => {},
    openRuntimeOptionsPalette: () => {},
    printHelp: () => {},
    stopApp: () => {},
    hasMainWindow: () => true,
    getMultiCopyTimeoutMs: () => 5000,
    schedule: (fn) => setTimeout(fn, 0),
    logInfo: () => {},
    logDebug: () => {},
    logWarn: () => {},
    logError: () => {},
  });

  const context = createContext();
  context.setSocketPath('/tmp/new.sock');
  context.showOsd('hello');
  context.setVisibleOverlay(true);
  context.toggleVisibleOverlay();

  assert.equal(appState.mpvSocketPath, '/tmp/new.sock');
  assert.deepEqual(calls, ['osd:hello', 'set-visible:true', 'toggle-visible']);
});

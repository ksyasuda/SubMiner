import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildCliCommandContextMainDepsHandler } from './cli-command-context-main-deps';

test('cli command context main deps builder maps state and callbacks', async () => {
  const calls: string[] = [];
  const appState = {
    mpvSocketPath: '/tmp/mpv.sock',
    mpvClient: null,
    texthookerPort: 5174,
    overlayRuntimeInitialized: false,
  };

  const build = createBuildCliCommandContextMainDepsHandler({
    appState,
    texthookerService: { isRunning: () => false, start: () => null },
    getResolvedConfig: () => ({
      texthooker: { openBrowser: true },
      annotationWebsocket: { enabled: true, port: 6678 },
    }),
    defaultWebsocketPort: 6677,
    defaultAnnotationWebsocketPort: 6678,
    hasMpvWebsocketPlugin: () => false,
    openExternal: async (url) => {
      calls.push(`open:${url}`);
    },
    logBrowserOpenError: (url) => calls.push(`open-error:${url}`),
    showMpvOsd: (text) => calls.push(`osd:${text}`),

    initializeOverlayRuntime: () => calls.push('init-overlay'),
    toggleVisibleOverlay: () => calls.push('toggle-visible'),
    togglePrimarySubtitleBar: () => calls.push('toggle-primary-subtitle'),
    openFirstRunSetupWindow: (force?: boolean) =>
      calls.push(`open-setup:${force === true ? 'force' : 'default'}`),
    setVisibleOverlayVisible: (visible) => calls.push(`set-visible:${visible}`),

    copyCurrentSubtitle: () => calls.push('copy-sub'),
    startPendingMultiCopy: (timeoutMs) => calls.push(`multi:${timeoutMs}`),
    mineSentenceCard: async () => {
      calls.push('mine');
    },
    startPendingMineSentenceMultiple: (timeoutMs) => calls.push(`mine-multi:${timeoutMs}`),
    updateLastCardFromClipboard: async () => {
      calls.push('update-last-card');
    },
    refreshKnownWordCache: async () => {
      calls.push('refresh-known');
    },
    triggerFieldGrouping: async () => {
      calls.push('field-grouping');
    },
    triggerSubsyncFromConfig: async () => {
      calls.push('subsync');
    },
    markLastCardAsAudioCard: async () => {
      calls.push('mark-audio');
    },
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
    clearAnilistToken: () => calls.push('clear-token'),
    openAnilistSetupWindow: () => calls.push('open-anilist-setup'),
    openJellyfinSetupWindow: () => calls.push('open-jellyfin-setup'),
    getAnilistQueueStatus: () => ({
      pending: 1,
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
    runStatsCommand: async () => {
      calls.push('run-stats');
    },
    runJellyfinCommand: async () => {
      calls.push('run-jellyfin');
    },
    runUpdateCommand: async () => {
      calls.push('run-update');
    },
    runYoutubePlaybackFlow: async () => {
      calls.push('run-youtube-playback');
    },
    openYomitanSettings: () => calls.push('open-yomitan'),
    cycleSecondarySubMode: () => calls.push('cycle-secondary'),
    openRuntimeOptionsPalette: () => calls.push('open-runtime-options'),
    printHelp: () => calls.push('help'),
    stopApp: () => calls.push('stop-app'),
    hasMainWindow: () => true,
    getMultiCopyTimeoutMs: () => 5000,
    schedule: (fn) => {
      fn();
      return setTimeout(() => {}, 0);
    },
    logInfo: (message) => calls.push(`info:${message}`),
    logDebug: (message) => calls.push(`debug:${message}`),
    logWarn: (message) => calls.push(`warn:${message}`),
    logError: (message) => calls.push(`error:${message}`),
  });

  const deps = build();
  assert.equal(deps.getSocketPath(), '/tmp/mpv.sock');
  deps.setSocketPath('/tmp/next.sock');
  assert.equal(appState.mpvSocketPath, '/tmp/next.sock');
  assert.equal(deps.getTexthookerPort(), 5174);
  deps.setTexthookerPort(5175);
  assert.equal(appState.texthookerPort, 5175);
  assert.equal(deps.getTexthookerWebsocketUrl(), 'ws://127.0.0.1:6678');
  assert.equal(deps.shouldOpenBrowser(), true);
  deps.showOsd('hello');
  deps.initializeOverlay();
  deps.openFirstRunSetup(true);
  deps.setVisibleOverlay(true);
  deps.printHelp();

  assert.deepEqual(calls, [
    'osd:hello',
    'init-overlay',
    'open-setup:force',
    'set-visible:true',
    'help',
  ]);

  const retry = await deps.retryAnilistQueueNow();
  assert.deepEqual(retry, { ok: true, message: 'ok' });
});

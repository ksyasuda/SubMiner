import test from 'node:test';
import assert from 'node:assert/strict';
import { createBuildCliCommandContextDepsHandler } from './cli-command-context-deps';

test('build cli command context deps maps handlers and values', () => {
  const calls: string[] = [];
  const buildDeps = createBuildCliCommandContextDepsHandler({
    getSocketPath: () => '/tmp/mpv.sock',
    setSocketPath: (socketPath) => calls.push(`socket:${socketPath}`),
    getMpvClient: () => null,
    showOsd: (text) => calls.push(`osd:${text}`),
    texthookerService: { start: () => null, status: () => ({ running: false }) } as never,
    getTexthookerPort: () => 5174,
    setTexthookerPort: (port) => calls.push(`port:${port}`),
    getTexthookerWebsocketUrl: () => 'ws://127.0.0.1:6678',
    shouldOpenBrowser: () => true,
    openExternal: async (url) => calls.push(`open:${url}`),
    logBrowserOpenError: (url) => calls.push(`open-error:${url}`),
    isOverlayInitialized: () => true,
    initializeOverlay: () => calls.push('init'),
    toggleVisibleOverlay: () => calls.push('toggle-visible'),
    togglePrimarySubtitleBar: () => calls.push('toggle-primary-subtitle'),
    openFirstRunSetup: () => calls.push('setup'),
    setVisibleOverlay: (visible) => calls.push(`set-visible:${visible}`),
    copyCurrentSubtitle: () => calls.push('copy'),
    startPendingMultiCopy: (ms) => calls.push(`multi:${ms}`),
    mineSentenceCard: async () => {
      calls.push('mine');
    },
    startPendingMineSentenceMultiple: (ms) => calls.push(`mine-multi:${ms}`),
    updateLastCardFromClipboard: async () => {
      calls.push('update');
    },
    refreshKnownWordCache: async () => {
      calls.push('refresh');
    },
    triggerFieldGrouping: async () => {
      calls.push('group');
    },
    triggerSubsyncFromConfig: async () => {
      calls.push('subsync');
    },
    markLastCardAsAudioCard: async () => {
      calls.push('mark');
    },
    dispatchSessionAction: async () => {},
    getAnilistStatus: () => ({}) as never,
    clearAnilistToken: () => calls.push('clear-token'),
    openAnilistSetup: () => calls.push('anilist'),
    openJellyfinSetup: () => calls.push('jellyfin'),
    getAnilistQueueStatus: () => ({}) as never,
    retryAnilistQueueNow: async () => ({ ok: true, message: 'ok' }),
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
    openYomitanSettings: () => calls.push('yomitan'),
    cycleSecondarySubMode: () => calls.push('cycle-secondary'),
    openRuntimeOptionsPalette: () => calls.push('runtime-options'),
    printHelp: () => calls.push('help'),
    stopApp: () => calls.push('stop'),
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

  const deps = buildDeps();
  assert.equal(deps.getSocketPath(), '/tmp/mpv.sock');
  assert.equal(deps.getTexthookerPort(), 5174);
  assert.equal(deps.getTexthookerWebsocketUrl(), 'ws://127.0.0.1:6678');
  assert.equal(deps.shouldOpenBrowser(), true);
  assert.equal(deps.isOverlayInitialized(), true);
  assert.equal(deps.hasMainWindow(), true);
  assert.equal(deps.getMultiCopyTimeoutMs(), 5000);

  deps.setSocketPath('/tmp/next.sock');
  deps.showOsd('hello');
  deps.setTexthookerPort(5175);
  deps.printHelp();
  assert.deepEqual(calls, ['socket:/tmp/next.sock', 'osd:hello', 'port:5175', 'help']);
});

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
    getResolvedConfig: () => ({ texthooker: { openBrowser: true } }),
    openExternal: async (url) => {
      calls.push(`open:${url}`);
    },
    logBrowserOpenError: (url) => calls.push(`open-error:${url}`),
    showMpvOsd: (text) => calls.push(`osd:${text}`),

    initializeOverlayRuntime: () => calls.push('init-overlay'),
    toggleVisibleOverlay: () => calls.push('toggle-visible'),
    toggleInvisibleOverlay: () => calls.push('toggle-invisible'),
    setVisibleOverlayVisible: (visible) => calls.push(`set-visible:${visible}`),
    setInvisibleOverlayVisible: (visible) => calls.push(`set-invisible:${visible}`),

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
    runJellyfinCommand: async () => {
      calls.push('run-jellyfin');
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
  assert.equal(deps.shouldOpenBrowser(), true);
  deps.showOsd('hello');
  deps.initializeOverlay();
  deps.setVisibleOverlay(true);
  deps.setInvisibleOverlay(false);
  deps.printHelp();

  assert.deepEqual(calls, [
    'osd:hello',
    'init-overlay',
    'set-visible:true',
    'set-invisible:false',
    'help',
  ]);

  const retry = await deps.retryAnilistQueueNow();
  assert.deepEqual(retry, { ok: true, message: 'ok' });
});

import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildCliCommandContextMainDepsHandler } from './cli-command-context-main-deps';

type MainDeps = Parameters<typeof createBuildCliCommandContextMainDepsHandler>[0];

function buildContextDeps(overrides: Partial<MainDeps> = {}) {
  const noop = () => {};
  const noopAsync = async () => {};
  return createBuildCliCommandContextMainDepsHandler({
    appState: {
      mpvSocketPath: '/tmp/mpv.sock',
      mpvClient: null,
      texthookerPort: 5174,
      overlayRuntimeInitialized: false,
    },
    texthookerService: { isRunning: () => false, start: () => null },
    getResolvedConfig: () => ({}),
    defaultWebsocketPort: 6677,
    defaultAnnotationWebsocketPort: 6678,
    hasMpvWebsocketPlugin: () => false,
    openExternal: noopAsync,
    logBrowserOpenError: noop,
    showMpvOsd: noop,
    initializeOverlayRuntime: noop,
    toggleVisibleOverlay: noop,
    togglePrimarySubtitleBar: noop,
    openFirstRunSetupWindow: noop,
    setVisibleOverlayVisible: noop,
    copyCurrentSubtitle: noop,
    startPendingMultiCopy: noop,
    mineSentenceCard: noopAsync,
    startPendingMineSentenceMultiple: noop,
    updateLastCardFromClipboard: noopAsync,
    refreshKnownWordCache: noopAsync,
    triggerFieldGrouping: noopAsync,
    triggerSubsyncFromConfig: noopAsync,
    markLastCardAsAudioCard: noopAsync,
    dispatchSessionAction: noopAsync,
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
    clearAnilistToken: noop,
    openAnilistSetupWindow: noop,
    openJellyfinSetupWindow: noop,
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
    runStatsCommand: noopAsync,
    runJellyfinCommand: noopAsync,
    runUpdateCommand: noopAsync,
    runEnsureLinuxRuntimePluginAssetsCommand: noopAsync,
    runYoutubePlaybackFlow: noopAsync,
    openYomitanSettings: noop,
    openConfigSettingsWindow: noop,
    openSyncUiWindow: noop,
    openYoutubeBrowserWindow: noop,
    cycleSecondarySubMode: noop,
    openRuntimeOptionsPalette: noop,
    printHelp: noop,
    stopApp: noop,
    hasMainWindow: () => true,
    getMultiCopyTimeoutMs: () => 5000,
    schedule: (fn) => setTimeout(fn, 0),
    logInfo: noop,
    logDebug: noop,
    logWarn: noop,
    logError: noop,
    ...overrides,
  })();
}

test('cli command context setSocketPath notifies only when the socket path changes', () => {
  const changes: string[] = [];
  const appState = {
    mpvSocketPath: '/tmp/mpv.sock',
    mpvClient: null,
    texthookerPort: 5174,
    overlayRuntimeInitialized: false,
  };
  const deps = buildContextDeps({
    appState,
    onMpvSocketPathChanged: (next, previous) => changes.push(`${previous}->${next}`),
  });

  deps.setSocketPath('/tmp/next.sock');
  deps.setSocketPath('/tmp/next.sock');

  assert.equal(appState.mpvSocketPath, '/tmp/next.sock');
  assert.equal(deps.getSocketPath(), '/tmp/next.sock');
  assert.deepEqual(changes, ['/tmp/mpv.sock->/tmp/next.sock']);
});

for (const c of [
  { openBrowser: undefined, expected: true },
  { openBrowser: true, expected: true },
  { openBrowser: false, expected: false },
]) {
  test(`cli command context shouldOpenBrowser is ${c.expected} when texthooker.openBrowser is ${c.openBrowser}`, () => {
    const deps = buildContextDeps({
      getResolvedConfig: () => ({ texthooker: { openBrowser: c.openBrowser } }),
    });
    assert.equal(deps.shouldOpenBrowser(), c.expected);
  });
}

for (const c of [
  {
    name: 'uses the annotation websocket port when annotations are enabled',
    config: { annotationWebsocket: { enabled: true } },
    hasPlugin: true,
    expected: 'ws://127.0.0.1:6678',
  },
  {
    name: 'falls back to the default websocket port without the mpv plugin',
    config: { annotationWebsocket: { enabled: false } },
    hasPlugin: false,
    expected: 'ws://127.0.0.1:6677',
  },
  {
    name: 'returns no websocket url when the mpv plugin owns the websocket',
    config: { annotationWebsocket: { enabled: false } },
    hasPlugin: true,
    expected: undefined,
  },
]) {
  test(`cli command context texthooker websocket url ${c.name}`, () => {
    const deps = buildContextDeps({
      getResolvedConfig: () => c.config,
      hasMpvWebsocketPlugin: () => c.hasPlugin,
    });
    assert.equal(deps.getTexthookerWebsocketUrl(), c.expected);
  });
}

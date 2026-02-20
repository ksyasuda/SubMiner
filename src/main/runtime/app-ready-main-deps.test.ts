import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildAppReadyRuntimeMainDepsHandler } from './app-ready-main-deps';

test('app-ready main deps builder returns mapped app-ready runtime deps', async () => {
  const calls: string[] = [];
  const onReady = createBuildAppReadyRuntimeMainDepsHandler({
    loadSubtitlePosition: () => calls.push('load-subtitle-position'),
    resolveKeybindings: () => calls.push('resolve-keybindings'),
    createMpvClient: () => calls.push('create-mpv-client'),
    reloadConfig: () => calls.push('reload-config'),
    getResolvedConfig: () => ({ websocket: {} }),
    getConfigWarnings: () => [],
    logConfigWarning: () => calls.push('log-config-warning'),
    initRuntimeOptionsManager: () => calls.push('init-runtime-options'),
    setSecondarySubMode: () => calls.push('set-secondary-sub-mode'),
    defaultSecondarySubMode: 'hover',
    defaultWebsocketPort: 5174,
    hasMpvWebsocketPlugin: () => false,
    startSubtitleWebsocket: () => calls.push('start-ws'),
    log: () => calls.push('log'),
    setLogLevel: () => calls.push('set-log-level'),
    createMecabTokenizerAndCheck: async () => {
      calls.push('create-mecab');
    },
    createSubtitleTimingTracker: () => calls.push('create-subtitle-tracker'),
    createImmersionTracker: () => calls.push('create-immersion'),
    startJellyfinRemoteSession: async () => {
      calls.push('start-jellyfin');
    },
    loadYomitanExtension: async () => {
      calls.push('load-yomitan');
    },
    prewarmSubtitleDictionaries: async () => {
      calls.push('prewarm-dicts');
    },
    startBackgroundWarmups: () => calls.push('start-warmups'),
    texthookerOnlyMode: false,
    shouldAutoInitializeOverlayRuntimeFromConfig: () => true,
    initializeOverlayRuntime: () => calls.push('init-overlay'),
    handleInitialArgs: () => calls.push('handle-initial-args'),
    onCriticalConfigErrors: () => {
      throw new Error('should not call');
    },
    logDebug: () => calls.push('debug'),
    now: () => 123,
  })();

  assert.equal(onReady.defaultSecondarySubMode, 'hover');
  assert.equal(onReady.defaultWebsocketPort, 5174);
  assert.equal(onReady.texthookerOnlyMode, false);
  assert.equal(onReady.shouldAutoInitializeOverlayRuntimeFromConfig(), true);
  assert.equal(onReady.now?.(), 123);
  onReady.loadSubtitlePosition();
  onReady.resolveKeybindings();
  onReady.createMpvClient();
  await onReady.createMecabTokenizerAndCheck();
  await onReady.loadYomitanExtension();
  await onReady.prewarmSubtitleDictionaries?.();
  onReady.startBackgroundWarmups();

  assert.deepEqual(calls, [
    'load-subtitle-position',
    'resolve-keybindings',
    'create-mpv-client',
    'create-mecab',
    'load-yomitan',
    'prewarm-dicts',
    'start-warmups',
  ]);
});

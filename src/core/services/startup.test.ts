import assert from 'node:assert/strict';
import test from 'node:test';
import { runAppReadyRuntime } from './startup';

test('runAppReadyRuntime minimal startup skips Yomitan and first-run setup while still handling CLI args', async () => {
  const calls: string[] = [];

  await runAppReadyRuntime({
    ensureDefaultConfigBootstrap: () => {
      calls.push('bootstrap');
    },
    loadSubtitlePosition: () => {
      calls.push('load-subtitle-position');
    },
    resolveKeybindings: () => {
      calls.push('resolve-keybindings');
    },
    createMpvClient: () => {
      calls.push('create-mpv');
    },
    reloadConfig: () => {
      calls.push('reload-config');
    },
    getResolvedConfig: () => ({}),
    getConfigWarnings: () => [],
    logConfigWarning: () => {
      calls.push('config-warning');
    },
    setLogLevel: () => {
      calls.push('set-log-level');
    },
    initRuntimeOptionsManager: () => {
      calls.push('init-runtime-options');
    },
    setSecondarySubMode: () => {
      calls.push('set-secondary-sub-mode');
    },
    defaultSecondarySubMode: 'hover',
    defaultWebsocketPort: 0,
    defaultAnnotationWebsocketPort: 0,
    defaultTexthookerPort: 0,
    hasMpvWebsocketPlugin: () => false,
    startSubtitleWebsocket: () => {
      calls.push('subtitle-ws');
    },
    startAnnotationWebsocket: () => {
      calls.push('annotation-ws');
    },
    startTexthooker: () => {
      calls.push('texthooker');
    },
    log: () => {
      calls.push('log');
    },
    createMecabTokenizerAndCheck: async () => {
      calls.push('mecab');
    },
    createSubtitleTimingTracker: () => {
      calls.push('subtitle-timing');
    },
    createImmersionTracker: () => {
      calls.push('immersion');
    },
    startJellyfinRemoteSession: async () => {
      calls.push('jellyfin');
    },
    loadYomitanExtension: async () => {
      calls.push('load-yomitan');
    },
    handleFirstRunSetup: async () => {
      calls.push('first-run');
    },
    prewarmSubtitleDictionaries: async () => {
      calls.push('prewarm');
    },
    startBackgroundWarmups: () => {
      calls.push('warmups');
    },
    texthookerOnlyMode: false,
    shouldAutoInitializeOverlayRuntimeFromConfig: () => false,
    setVisibleOverlayVisible: () => {
      calls.push('visible-overlay');
    },
    initializeOverlayRuntime: () => {
      calls.push('init-overlay');
    },
    handleInitialArgs: () => {
      calls.push('handle-initial-args');
    },
    shouldUseMinimalStartup: () => true,
    shouldSkipHeavyStartup: () => false,
  });

  assert.deepEqual(calls, ['bootstrap', 'reload-config', 'handle-initial-args']);
});

test('runAppReadyRuntime headless refresh bootstraps Anki runtime without UI startup', async () => {
  const calls: string[] = [];

  await runAppReadyRuntime({
    ensureDefaultConfigBootstrap: () => {
      calls.push('bootstrap');
    },
    loadSubtitlePosition: () => {
      calls.push('load-subtitle-position');
    },
    resolveKeybindings: () => {
      calls.push('resolve-keybindings');
    },
    createMpvClient: () => {
      calls.push('create-mpv');
    },
    reloadConfig: () => {
      calls.push('reload-config');
    },
    getResolvedConfig: () => ({}),
    getConfigWarnings: () => [],
    logConfigWarning: () => {
      calls.push('config-warning');
    },
    setLogLevel: () => {
      calls.push('set-log-level');
    },
    initRuntimeOptionsManager: () => {
      calls.push('init-runtime-options');
    },
    setSecondarySubMode: () => {
      calls.push('set-secondary-sub-mode');
    },
    defaultSecondarySubMode: 'hover',
    defaultWebsocketPort: 0,
    defaultAnnotationWebsocketPort: 0,
    defaultTexthookerPort: 0,
    hasMpvWebsocketPlugin: () => false,
    startSubtitleWebsocket: () => {
      calls.push('subtitle-ws');
    },
    startAnnotationWebsocket: () => {
      calls.push('annotation-ws');
    },
    startTexthooker: () => {
      calls.push('texthooker');
    },
    log: () => {
      calls.push('log');
    },
    createMecabTokenizerAndCheck: async () => {
      calls.push('mecab');
    },
    createSubtitleTimingTracker: () => {
      calls.push('subtitle-timing');
    },
    createImmersionTracker: () => {
      calls.push('immersion');
    },
    startJellyfinRemoteSession: async () => {
      calls.push('jellyfin');
    },
    loadYomitanExtension: async () => {
      calls.push('load-yomitan');
    },
    handleFirstRunSetup: async () => {
      calls.push('first-run');
    },
    prewarmSubtitleDictionaries: async () => {
      calls.push('prewarm');
    },
    startBackgroundWarmups: () => {
      calls.push('warmups');
    },
    texthookerOnlyMode: false,
    shouldAutoInitializeOverlayRuntimeFromConfig: () => false,
    setVisibleOverlayVisible: () => {
      calls.push('visible-overlay');
    },
    initializeOverlayRuntime: () => {
      calls.push('init-overlay');
    },
    runHeadlessInitialCommand: async () => {
      calls.push('run-headless-command');
    },
    handleInitialArgs: () => {
      calls.push('handle-initial-args');
    },
    shouldRunHeadlessInitialCommand: () => true,
    shouldUseMinimalStartup: () => false,
    shouldSkipHeavyStartup: () => false,
  });

  assert.deepEqual(calls, [
    'bootstrap',
    'reload-config',
    'init-runtime-options',
    'run-headless-command',
  ]);
});

test('runAppReadyRuntime loads Yomitan before headless overlay fallback initialization', async () => {
  const calls: string[] = [];

  await runAppReadyRuntime({
    ensureDefaultConfigBootstrap: () => {
      calls.push('bootstrap');
    },
    loadSubtitlePosition: () => {
      calls.push('load-subtitle-position');
    },
    resolveKeybindings: () => {
      calls.push('resolve-keybindings');
    },
    createMpvClient: () => {
      calls.push('create-mpv');
    },
    reloadConfig: () => {
      calls.push('reload-config');
    },
    getResolvedConfig: () => ({}),
    getConfigWarnings: () => [],
    logConfigWarning: () => {},
    setLogLevel: () => {},
    initRuntimeOptionsManager: () => {
      calls.push('init-runtime-options');
    },
    setSecondarySubMode: () => {},
    defaultSecondarySubMode: 'hover',
    defaultWebsocketPort: 0,
    defaultAnnotationWebsocketPort: 0,
    defaultTexthookerPort: 0,
    hasMpvWebsocketPlugin: () => false,
    startSubtitleWebsocket: () => {},
    startAnnotationWebsocket: () => {},
    startTexthooker: () => {},
    log: () => {},
    createMecabTokenizerAndCheck: async () => {},
    createSubtitleTimingTracker: () => {
      calls.push('subtitle-timing');
    },
    createImmersionTracker: () => {},
    startJellyfinRemoteSession: async () => {},
    loadYomitanExtension: async () => {
      calls.push('load-yomitan');
    },
    handleFirstRunSetup: async () => {},
    prewarmSubtitleDictionaries: async () => {},
    startBackgroundWarmups: () => {},
    texthookerOnlyMode: false,
    shouldAutoInitializeOverlayRuntimeFromConfig: () => false,
    setVisibleOverlayVisible: () => {},
    initializeOverlayRuntime: () => {
      calls.push('init-overlay');
    },
    handleInitialArgs: () => {
      calls.push('handle-initial-args');
    },
    shouldRunHeadlessInitialCommand: () => true,
    shouldUseMinimalStartup: () => false,
    shouldSkipHeavyStartup: () => false,
  });

  assert.deepEqual(calls, [
    'bootstrap',
    'reload-config',
    'init-runtime-options',
    'create-mpv',
    'subtitle-timing',
    'load-yomitan',
    'init-overlay',
    'handle-initial-args',
  ]);
});

test('runAppReadyRuntime loads Yomitan before auto-initializing overlay runtime', async () => {
  const calls: string[] = [];

  await runAppReadyRuntime({
    ensureDefaultConfigBootstrap: () => {
      calls.push('bootstrap');
    },
    loadSubtitlePosition: () => {
      calls.push('load-subtitle-position');
    },
    resolveKeybindings: () => {
      calls.push('resolve-keybindings');
    },
    createMpvClient: () => {
      calls.push('create-mpv');
    },
    reloadConfig: () => {
      calls.push('reload-config');
    },
    getResolvedConfig: () => ({
      websocket: { enabled: false },
      annotationWebsocket: { enabled: false },
      texthooker: { launchAtStartup: false },
    }),
    getConfigWarnings: () => [],
    logConfigWarning: () => {},
    setLogLevel: () => {
      calls.push('set-log-level');
    },
    initRuntimeOptionsManager: () => {
      calls.push('init-runtime-options');
    },
    setSecondarySubMode: () => {
      calls.push('set-secondary-sub-mode');
    },
    defaultSecondarySubMode: 'hover',
    defaultWebsocketPort: 0,
    defaultAnnotationWebsocketPort: 0,
    defaultTexthookerPort: 0,
    hasMpvWebsocketPlugin: () => false,
    startSubtitleWebsocket: () => {
      calls.push('subtitle-ws');
    },
    startAnnotationWebsocket: () => {
      calls.push('annotation-ws');
    },
    startTexthooker: () => {
      calls.push('texthooker');
    },
    log: () => {
      calls.push('log');
    },
    createMecabTokenizerAndCheck: async () => {},
    createSubtitleTimingTracker: () => {
      calls.push('subtitle-timing');
    },
    createImmersionTracker: () => {
      calls.push('immersion');
    },
    startJellyfinRemoteSession: async () => {},
    loadYomitanExtension: async () => {
      calls.push('load-yomitan');
    },
    handleFirstRunSetup: async () => {
      calls.push('first-run');
    },
    prewarmSubtitleDictionaries: async () => {},
    startBackgroundWarmups: () => {
      calls.push('warmups');
    },
    texthookerOnlyMode: false,
    shouldAutoInitializeOverlayRuntimeFromConfig: () => true,
    setVisibleOverlayVisible: () => {
      calls.push('visible-overlay');
    },
    initializeOverlayRuntime: () => {
      calls.push('init-overlay');
    },
    handleInitialArgs: () => {
      calls.push('handle-initial-args');
    },
    shouldUseMinimalStartup: () => false,
    shouldSkipHeavyStartup: () => false,
  });

  assert.ok(calls.indexOf('load-yomitan') !== -1);
  assert.ok(calls.indexOf('init-overlay') !== -1);
  assert.ok(calls.indexOf('load-yomitan') < calls.indexOf('init-overlay'));
});

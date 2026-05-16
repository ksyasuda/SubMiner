import assert from 'node:assert/strict';
import test from 'node:test';
import { composeAppReadyRuntime } from './app-ready-composer';

test('composeAppReadyRuntime returns reload/critical/app-ready handlers', () => {
  const calls: string[] = [];
  const composed = composeAppReadyRuntime({
    reloadConfigMainDeps: {
      reloadConfigStrict: () => {
        calls.push('reloadConfigStrict');
        return { ok: true, path: '/tmp/config.jsonc', warnings: [] };
      },
      logInfo: () => {},
      logDebug: () => {},
      logWarning: () => {},
      showDesktopNotification: () => {},
      startConfigHotReload: () => {},
      refreshAnilistClientSecretState: async () => {},
      failHandlers: {
        logError: () => {},
        showErrorBox: () => {},
        quit: () => {},
      },
    },
    criticalConfigErrorMainDeps: {
      getConfigPath: () => '/tmp/config.jsonc',
      failHandlers: {
        logError: () => {},
        showErrorBox: () => {},
        quit: () => {},
      },
    },
    appReadyRuntimeMainDeps: {
      ensureDefaultConfigBootstrap: () => {},
      loadSubtitlePosition: () => {},
      resolveKeybindings: () => {},
      createMpvClient: () => {},
      getResolvedConfig: () => ({}) as never,
      getConfigWarnings: () => [],
      logConfigWarning: () => {},
      setLogLevel: () => {},
      initRuntimeOptionsManager: () => {},
      setSecondarySubMode: () => {},
      defaultSecondarySubMode: 'hover',
      defaultWebsocketPort: 5174,
      defaultAnnotationWebsocketPort: 6678,
      defaultTexthookerPort: 5174,
      hasMpvWebsocketPlugin: () => false,
      startSubtitleWebsocket: () => {},
      startAnnotationWebsocket: () => {},
      startTexthooker: () => {},
      log: () => {},
      createMecabTokenizerAndCheck: async () => {},
      createSubtitleTimingTracker: () => {},
      loadYomitanExtension: async () => {},
      handleFirstRunSetup: async () => {},
      startJellyfinRemoteSession: async () => {},
      prewarmSubtitleDictionaries: async () => {},
      startBackgroundWarmups: () => {},
      texthookerOnlyMode: false,
      shouldAutoInitializeOverlayRuntimeFromConfig: () => false,
      setVisibleOverlayVisible: () => {},
      initializeOverlayRuntime: () => {},
      handleInitialArgs: () => {},
      logDebug: () => {},
      now: () => Date.now(),
    },
    immersionTrackerStartupMainDeps: {
      getResolvedConfig: () => ({}) as never,
      getConfiguredDbPath: () => '/tmp/immersion.sqlite',
      createTrackerService: () =>
        ({
          startSession: () => {},
        }) as never,
      setTracker: () => {},
      getMpvClient: () => null,
      seedTrackerFromCurrentMedia: () => {},
      logInfo: () => {},
      logDebug: () => {},
      logWarn: () => {},
    },
  });

  assert.equal(typeof composed.reloadConfig, 'function');
  assert.equal(typeof composed.criticalConfigError, 'function');
  assert.equal(typeof composed.appReadyRuntimeRunner, 'function');

  // reloadConfig invokes the injected reloadConfigStrict dep
  composed.reloadConfig();
  assert.deepEqual(calls, ['reloadConfigStrict']);
});

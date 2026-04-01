import assert from 'node:assert/strict';
import test from 'node:test';

import { createAppReadyRuntime } from './app-ready-runtime';

test('app ready runtime shares overlay startup prereqs with youtube runtime init path', async () => {
  let subtitlePosition: unknown | null = null;
  let keybindingsCount = 0;
  let hasMpvClient = false;
  let runtimeOptionsManager: unknown | null = null;
  let subtitleTimingTracker: unknown | null = null;
  let overlayRuntimeInitialized = false;
  const calls: string[] = [];

  const runtime = createAppReadyRuntime({
    reload: {
      reloadConfigStrict: () => ({ ok: true as const, path: '/tmp/config.jsonc', warnings: [] }),
      logInfo: () => {},
      logWarning: () => {},
      showDesktopNotification: () => {},
      startConfigHotReload: () => {},
      refreshAnilistClientSecretState: async () => undefined,
      failHandlers: {
        logError: () => {},
        showErrorBox: () => {},
        quit: () => {},
      },
    },
    criticalConfig: {
      getConfigPath: () => '/tmp/config.jsonc',
      failHandlers: {
        logError: () => {},
        showErrorBox: () => {},
        quit: () => {
          throw new Error('quit');
        },
      },
    },
    immersion: {
      getResolvedConfig: () =>
        ({
          immersionTracking: {
            enabled: true,
            batchSize: 1,
            flushIntervalMs: 1,
            queueCap: 1,
            payloadCapBytes: 1,
            maintenanceIntervalMs: 1,
            retention: {
              eventsDays: 1,
              telemetryDays: 1,
              sessionsDays: 1,
              dailyRollupsDays: 1,
              monthlyRollupsDays: 1,
              vacuumIntervalDays: 1,
            },
          },
        }) as never,
      getConfiguredDbPath: () => '/tmp/immersion.sqlite',
      createTrackerService: () => ({}) as never,
      setTracker: () => {},
      getMpvClient: () => null,
      seedTrackerFromCurrentMedia: () => {},
      logInfo: () => {},
      logDebug: () => {},
      logWarn: () => {},
    },
    runner: {
      ensureDefaultConfigBootstrap: () => {},
      getSubtitlePosition: () => subtitlePosition,
      loadSubtitlePosition: () => {
        subtitlePosition = { mode: 'bottom' };
        calls.push('loadSubtitlePosition');
      },
      getKeybindingsCount: () => keybindingsCount,
      resolveKeybindings: () => {
        keybindingsCount = 3;
        calls.push('resolveKeybindings');
      },
      hasMpvClient: () => hasMpvClient,
      createMpvClient: () => {
        hasMpvClient = true;
        calls.push('createMpvClient');
      },
      getRuntimeOptionsManager: () => runtimeOptionsManager,
      initRuntimeOptionsManager: () => {
        runtimeOptionsManager = {};
        calls.push('initRuntimeOptionsManager');
      },
      getSubtitleTimingTracker: () => subtitleTimingTracker,
      createSubtitleTimingTracker: () => {
        subtitleTimingTracker = {};
        calls.push('createSubtitleTimingTracker');
      },
      getResolvedConfig: () =>
        ({
          ankiConnect: {
            enabled: false,
          },
        }) as never,
      getConfigWarnings: () => [],
      logConfigWarning: () => {},
      setLogLevel: () => {},
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
      loadYomitanExtension: async () => {},
      ensureYomitanExtensionLoaded: async () => {
        calls.push('ensureYomitanExtensionLoaded');
      },
      handleFirstRunSetup: async () => {},
      startBackgroundWarmups: () => {},
      texthookerOnlyMode: false,
      shouldAutoInitializeOverlayRuntimeFromConfig: () => false,
      setVisibleOverlayVisible: () => {},
      initializeOverlayRuntime: () => {
        overlayRuntimeInitialized = true;
        calls.push('initializeOverlayRuntime');
      },
      ensureOverlayWindowsReadyForVisibilityActions: () => {
        calls.push('ensureOverlayWindowsReadyForVisibilityActions');
      },
      handleInitialArgs: () => {},
    },
    isOverlayRuntimeInitialized: () => overlayRuntimeInitialized,
  });

  runtime.ensureOverlayStartupPrereqs();
  runtime.ensureOverlayStartupPrereqs();
  await runtime.ensureYoutubePlaybackRuntimeReady();

  assert.deepEqual(calls, [
    'loadSubtitlePosition',
    'resolveKeybindings',
    'createMpvClient',
    'initRuntimeOptionsManager',
    'createSubtitleTimingTracker',
    'ensureYomitanExtensionLoaded',
    'initializeOverlayRuntime',
  ]);
});

test('app ready runtime reuses existing overlay runtime during youtube readiness', async () => {
  const calls: string[] = [];

  const runtime = createAppReadyRuntime({
    reload: {
      reloadConfigStrict: () => ({ ok: true as const, path: '/tmp/config.jsonc', warnings: [] }),
      logInfo: () => {},
      logWarning: () => {},
      showDesktopNotification: () => {},
      startConfigHotReload: () => {},
      refreshAnilistClientSecretState: async () => undefined,
      failHandlers: {
        logError: () => {},
        showErrorBox: () => {},
        quit: () => {},
      },
    },
    criticalConfig: {
      getConfigPath: () => '/tmp/config.jsonc',
      failHandlers: {
        logError: () => {},
        showErrorBox: () => {},
        quit: () => {
          throw new Error('quit');
        },
      },
    },
    immersion: {
      getResolvedConfig: () =>
        ({
          immersionTracking: {
            enabled: true,
            batchSize: 1,
            flushIntervalMs: 1,
            queueCap: 1,
            payloadCapBytes: 1,
            maintenanceIntervalMs: 1,
            retention: {
              eventsDays: 1,
              telemetryDays: 1,
              sessionsDays: 1,
              dailyRollupsDays: 1,
              monthlyRollupsDays: 1,
              vacuumIntervalDays: 1,
            },
          },
        }) as never,
      getConfiguredDbPath: () => '/tmp/immersion.sqlite',
      createTrackerService: () => ({}) as never,
      setTracker: () => {},
      getMpvClient: () => null,
      seedTrackerFromCurrentMedia: () => {},
      logInfo: () => {},
      logDebug: () => {},
      logWarn: () => {},
    },
    runner: {
      ensureDefaultConfigBootstrap: () => {},
      getSubtitlePosition: () => ({}) as never,
      loadSubtitlePosition: () => {
        throw new Error('should not load subtitle position');
      },
      getKeybindingsCount: () => 1,
      resolveKeybindings: () => {
        throw new Error('should not resolve keybindings');
      },
      hasMpvClient: () => true,
      createMpvClient: () => {
        throw new Error('should not create mpv client');
      },
      getRuntimeOptionsManager: () => ({}),
      initRuntimeOptionsManager: () => {
        throw new Error('should not init runtime options');
      },
      getSubtitleTimingTracker: () => ({}),
      createSubtitleTimingTracker: () => {
        throw new Error('should not create subtitle timing tracker');
      },
      getResolvedConfig: () => ({}) as never,
      getConfigWarnings: () => [],
      logConfigWarning: () => {},
      setLogLevel: () => {},
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
      loadYomitanExtension: async () => {},
      ensureYomitanExtensionLoaded: async () => {
        calls.push('ensureYomitanExtensionLoaded');
      },
      handleFirstRunSetup: async () => {},
      startBackgroundWarmups: () => {},
      texthookerOnlyMode: false,
      shouldAutoInitializeOverlayRuntimeFromConfig: () => false,
      setVisibleOverlayVisible: () => {},
      initializeOverlayRuntime: () => {
        calls.push('initializeOverlayRuntime');
      },
      ensureOverlayWindowsReadyForVisibilityActions: () => {
        calls.push('ensureOverlayWindowsReadyForVisibilityActions');
      },
      handleInitialArgs: () => {},
    },
    isOverlayRuntimeInitialized: () => true,
  });

  await runtime.ensureYoutubePlaybackRuntimeReady();

  assert.deepEqual(calls, [
    'ensureYomitanExtensionLoaded',
    'ensureOverlayWindowsReadyForVisibilityActions',
  ]);
});

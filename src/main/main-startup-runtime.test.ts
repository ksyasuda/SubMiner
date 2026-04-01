import assert from 'node:assert/strict';
import test from 'node:test';

import { createMainStartupRuntime } from './main-startup-runtime';

test('main startup runtime composes app-ready, cli, and headless runtimes', async () => {
  const calls: string[] = [];

  const runtime = createMainStartupRuntime<{ mode: string }>({
    appReady: {
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
          quit: () => {},
        },
      },
      immersion: {
        getResolvedConfig: () => ({ immersionTracking: { enabled: false } }) as never,
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
        ensureDefaultConfigBootstrap: () => {
          calls.push('ensureDefaultConfigBootstrap');
        },
        getSubtitlePosition: () => null,
        loadSubtitlePosition: () => {
          calls.push('loadSubtitlePosition');
        },
        getKeybindingsCount: () => 0,
        resolveKeybindings: () => {
          calls.push('resolveKeybindings');
        },
        hasMpvClient: () => false,
        createMpvClient: () => {
          calls.push('createMpvClient');
        },
        getRuntimeOptionsManager: () => null,
        initRuntimeOptionsManager: () => {
          calls.push('initRuntimeOptionsManager');
        },
        getSubtitleTimingTracker: () => null,
        createSubtitleTimingTracker: () => {
          calls.push('createSubtitleTimingTracker');
        },
        getResolvedConfig: () => ({ ankiConnect: { enabled: false } }) as never,
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
        startBackgroundWarmups: () => {
          calls.push('startBackgroundWarmups');
        },
        texthookerOnlyMode: false,
        shouldAutoInitializeOverlayRuntimeFromConfig: () => false,
        setVisibleOverlayVisible: () => {},
        initializeOverlayRuntime: () => {
          calls.push('initializeOverlayRuntime');
        },
        ensureOverlayWindowsReadyForVisibilityActions: () => {
          calls.push('ensureOverlayWindowsReadyForVisibilityActions');
        },
        handleInitialArgs: () => {
          calls.push('handleInitialArgs');
        },
      },
      isOverlayRuntimeInitialized: () => false,
    },
    cli: {
      appState: {
        appState: {} as never,
        getInitialArgs: () => null,
        isBackgroundMode: () => false,
        isTexthookerOnlyMode: () => false,
        setTexthookerOnlyMode: () => {},
        hasImmersionTracker: () => false,
        getMpvClient: () => null,
        isOverlayRuntimeInitialized: () => false,
      },
      config: {
        defaultConfig: { websocket: { port: 6677 }, annotationWebsocket: { port: 6678 } } as never,
        getResolvedConfig: () => ({}) as never,
        setCliLogLevel: () => {},
        hasMpvWebsocketPlugin: () => false,
      },
      io: {
        texthookerService: {} as never,
        openExternal: async () => {},
        logBrowserOpenError: () => {},
        showMpvOsd: () => {},
        schedule: () => 0 as never,
        logInfo: () => {},
        logWarn: () => {},
        logError: () => {},
      },
      commands: {
        initializeOverlayRuntime: () => {},
        toggleVisibleOverlay: () => {},
        openFirstRunSetupWindow: () => {},
        setVisibleOverlayVisible: () => {},
        copyCurrentSubtitle: () => {},
        startPendingMultiCopy: () => {},
        mineSentenceCard: async () => {},
        startPendingMineSentenceMultiple: () => {},
        updateLastCardFromClipboard: async () => {},
        refreshKnownWordCache: async () => {},
        triggerFieldGrouping: async () => {},
        triggerSubsyncFromConfig: async () => {},
        markLastCardAsAudioCard: async () => {},
        getAnilistStatus: () => ({}) as never,
        clearAnilistToken: () => {},
        openAnilistSetupWindow: () => {},
        openJellyfinSetupWindow: () => {},
        getAnilistQueueStatus: () => ({}) as never,
        processNextAnilistRetryUpdate: async () => ({ ok: true, message: 'done' }),
        generateCharacterDictionary: async () => ({
          zipPath: '/tmp/test.zip',
          fromCache: false,
          mediaId: 1,
          mediaTitle: 'Test',
          entryCount: 1,
        }),
        runJellyfinCommand: async () => {},
        runStatsCommand: async () => {},
        runYoutubePlaybackFlow: async () => {},
        openYomitanSettings: () => {},
        cycleSecondarySubMode: () => {},
        openRuntimeOptionsPalette: () => {},
        printHelp: () => {},
        stopApp: () => {
          calls.push('stopApp');
        },
        hasMainWindow: () => false,
        getMultiCopyTimeoutMs: () => 0,
      },
      startup: {
        shouldEnsureTrayOnStartup: () => false,
        shouldRunHeadlessInitialCommand: () => false,
        ensureTray: () => {},
        commandNeedsOverlayStartupPrereqs: () => false,
        commandNeedsOverlayRuntime: () => false,
        ensureOverlayStartupPrereqs: () => {
          calls.push('ensureOverlayStartupPrereqs');
        },
        startBackgroundWarmups: () => {
          calls.push('startupStartBackgroundWarmups');
        },
      },
      handleCliCommandRuntimeServiceWithContext: (args) => {
        calls.push(`handle:${(args as { command?: string }).command ?? 'unknown'}`);
      },
    },
    headless: {
      appLifecycleRuntimeRunnerMainDeps: {
        app: { on: () => {} } as never,
        platform: 'darwin',
        shouldStartApp: () => true,
        parseArgs: () => ({}) as never,
        handleCliCommand: () => {},
        printHelp: () => {},
        logNoRunningInstance: () => {},
        onReady: async () => {},
        onWillQuitCleanup: () => {},
        shouldRestoreWindowsOnActivate: () => false,
        restoreWindowsOnActivate: () => {},
        shouldQuitOnWindowAllClosed: () => false,
      },
      createAppLifecycleRuntimeRunner: () => (args) => {
        calls.push(`lifecycle:${(args as { command?: string }).command ?? 'unknown'}`);
      },
      bootstrap: {
        argv: ['node', 'main.js'],
        parseArgs: () => ({ command: 'start' }) as never,
        setLogLevel: () => {},
        forceX11Backend: () => {},
        enforceUnsupportedWaylandMode: () => {},
        shouldStartApp: () => true,
        getDefaultSocketPath: () => '/tmp/mpv.sock',
        defaultTexthookerPort: 5174,
        configDir: '/tmp/config',
        defaultConfig: {} as never,
        generateConfigTemplate: () => 'template',
        generateDefaultConfigFile: async () => 0,
        setExitCode: () => {},
        quitApp: () => {},
        logGenerateConfigError: () => {},
        startAppLifecycle: () => {},
      },
      runStartupBootstrapRuntime: (deps) => {
        deps.startAppLifecycle({ command: 'start' } as never);
        return { mode: 'started' };
      },
      applyStartupState: (state: { mode: string }) => {
        calls.push(`apply:${state.mode}`);
      },
    },
  });

  assert.equal(typeof runtime.appReady.runAppReady, 'function');
  assert.equal(typeof runtime.cliStartup.handleCliCommand, 'function');
  assert.equal(typeof runtime.headlessStartup.runAndApplyStartupState, 'function');
  assert.equal(runtime.handleCliCommand, runtime.cliStartup.handleCliCommand);
  assert.equal(runtime.handleInitialArgs, runtime.cliStartup.handleInitialArgs);
  assert.equal(
    runtime.appLifecycleRuntimeRunner,
    runtime.headlessStartup.appLifecycleRuntimeRunner,
  );
  assert.equal(runtime.runAndApplyStartupState, runtime.headlessStartup.runAndApplyStartupState);

  runtime.appReady.ensureOverlayStartupPrereqs();
  runtime.handleCliCommand({ command: 'start' } as never);
  assert.deepEqual(runtime.runAndApplyStartupState(), { mode: 'started' });

  assert.deepEqual(calls, [
    'loadSubtitlePosition',
    'resolveKeybindings',
    'createMpvClient',
    'initRuntimeOptionsManager',
    'createSubtitleTimingTracker',
    'handle:start',
    'lifecycle:start',
    'apply:started',
  ]);
});

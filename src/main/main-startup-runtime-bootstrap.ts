import type { MainStartupBootstrapInput } from './main-startup-bootstrap';
import type { MainStartupRuntime } from './main-startup-runtime';
import type { FirstRunRuntime } from './first-run-runtime';
import { createMainStartupBootstrap } from './main-startup-bootstrap';

type StartupBootstrapYomitanRuntime<TStartupState> = {
  loadYomitanExtension: MainStartupBootstrapInput<TStartupState>['runtime']['yomitan']['loadYomitanExtension'];
  ensureYomitanExtensionLoaded: MainStartupBootstrapInput<TStartupState>['runtime']['yomitan']['ensureYomitanExtensionLoaded'];
  openYomitanSettings: MainStartupBootstrapInput<TStartupState>['runtime']['yomitan']['openYomitanSettings'];
};

export interface MainStartupRuntimeBootstrapInput<TStartupState> {
  appState: MainStartupBootstrapInput<TStartupState>['appState'];
  appLifecycle: {
    app: MainStartupBootstrapInput<TStartupState>['appLifecycle']['app'];
    argv: string[];
    platform: NodeJS.Platform;
  };
  config: MainStartupBootstrapInput<TStartupState>['config'];
  logging: MainStartupBootstrapInput<TStartupState>['logging'];
  shell: MainStartupBootstrapInput<TStartupState>['shell'];
  runtime: Omit<MainStartupBootstrapInput<TStartupState>['runtime'], 'overlayUi' | 'yomitan'> & {
    texthookerService: {
      isRunning: () => boolean;
      start: (port: number, websocketUrl?: string) => void;
    };
    getOverlayUi: MainStartupBootstrapInput<TStartupState>['runtime']['overlayUi']['get'];
    getYomitanRuntime: () => StartupBootstrapYomitanRuntime<TStartupState>;
    getCharacterDictionaryDisabledReason: () => string | null;
  };
  commands: Omit<
    MainStartupBootstrapInput<TStartupState>['commands'],
    | 'startTexthooker'
    | 'generateCharacterDictionary'
    | 'runYoutubePlaybackFlow'
    | 'getMultiCopyTimeoutMs'
  > & {
    getConfiguredShortcuts: () => { multiCopyTimeoutMs: number };
    runYoutubePlaybackFlow: MainStartupBootstrapInput<TStartupState>['commands']['runYoutubePlaybackFlow'];
  };
  constants: MainStartupBootstrapInput<TStartupState>['constants'];
}

export interface MainStartupRuntimeBootstrap<TStartupState> {
  startupRuntime: MainStartupRuntime<TStartupState>;
}

export interface MainStartupRuntimeFromMainStateInput<TStartupState> {
  appState: MainStartupRuntimeBootstrapInput<TStartupState>['appState'];
  appLifecycle: MainStartupRuntimeBootstrapInput<TStartupState>['appLifecycle'];
  config: MainStartupRuntimeBootstrapInput<TStartupState>['config'];
  logging: MainStartupRuntimeBootstrapInput<TStartupState>['logging'];
  shell: MainStartupRuntimeBootstrapInput<TStartupState>['shell'];
  runtime: {
    subtitle: MainStartupRuntimeBootstrapInput<TStartupState>['runtime']['subtitle'];
    getOverlayUi: MainStartupRuntimeBootstrapInput<TStartupState>['runtime']['getOverlayUi'];
    overlayManager: MainStartupRuntimeBootstrapInput<TStartupState>['runtime']['overlayManager'];
    firstRun: {
      ensureSetupStateInitialized: FirstRunRuntime['ensureSetupStateInitialized'];
      openFirstRunSetupWindow: () => void;
    };
    anilist: {
      refreshAnilistClientSecretStateIfEnabled: MainStartupRuntimeBootstrapInput<TStartupState>['runtime']['anilist']['refreshAnilistClientSecretStateIfEnabled'];
      openAnilistSetupWindow: MainStartupRuntimeBootstrapInput<TStartupState>['runtime']['anilist']['openAnilistSetupWindow'];
      getStatusSnapshot: MainStartupRuntimeBootstrapInput<TStartupState>['runtime']['anilist']['getStatusSnapshot'];
      clearTokenState: MainStartupRuntimeBootstrapInput<TStartupState>['runtime']['anilist']['clearTokenState'];
      getQueueStatusSnapshot: MainStartupRuntimeBootstrapInput<TStartupState>['runtime']['anilist']['getQueueStatusSnapshot'];
      processNextAnilistRetryUpdate: MainStartupRuntimeBootstrapInput<TStartupState>['runtime']['anilist']['processNextAnilistRetryUpdate'];
    };
    jellyfin: MainStartupRuntimeBootstrapInput<TStartupState>['runtime']['jellyfin'];
    stats: MainStartupRuntimeBootstrapInput<TStartupState>['runtime']['stats'];
    mining: MainStartupRuntimeBootstrapInput<TStartupState>['runtime']['mining'];
    texthookerService: MainStartupRuntimeBootstrapInput<TStartupState>['runtime']['texthookerService'];
    yomitan: StartupBootstrapYomitanRuntime<TStartupState>;
    getCharacterDictionaryDisabledReason: () => string | null;
    subsyncRuntime: MainStartupRuntimeBootstrapInput<TStartupState>['runtime']['subsyncRuntime'];
    dictionarySupport: MainStartupRuntimeBootstrapInput<TStartupState>['runtime']['dictionarySupport'];
    subtitleWsService: MainStartupRuntimeBootstrapInput<TStartupState>['runtime']['subtitleWsService'];
    annotationSubtitleWsService: MainStartupRuntimeBootstrapInput<TStartupState>['runtime']['annotationSubtitleWsService'];
    immersion: MainStartupRuntimeBootstrapInput<TStartupState>['runtime']['immersion'];
  };
  commands: {
    mpvRuntime: {
      createMpvClientRuntimeService: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['createMpvClientRuntimeService'];
      createMecabTokenizerAndCheck: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['createMecabTokenizerAndCheck'];
      prewarmSubtitleDictionaries: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['prewarmSubtitleDictionaries'];
      startBackgroundWarmups: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['startBackgroundWarmups'];
    };
    runHeadlessInitialCommand: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['runHeadlessInitialCommand'];
    shortcuts: {
      startPendingMultiCopy: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['startPendingMultiCopy'];
      startPendingMineSentenceMultiple: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['startPendingMineSentenceMultiple'];
      refreshOverlayShortcuts: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['refreshOverlayShortcuts'];
      getConfiguredShortcuts: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['getConfiguredShortcuts'];
    };
    cycleSecondarySubMode: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['cycleSecondarySubMode'];
    hasMpvWebsocketPlugin: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['hasMpvWebsocketPlugin'];
    showMpvOsd: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['showMpvOsd'];
    shouldAutoOpenFirstRunSetup: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['shouldAutoOpenFirstRunSetup'];
    youtube: {
      runYoutubePlaybackFlow: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['runYoutubePlaybackFlow'];
    };
    shouldEnsureTrayOnStartupForInitialArgs: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['shouldEnsureTrayOnStartupForInitialArgs'];
    isHeadlessInitialCommand: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['isHeadlessInitialCommand'];
    commandNeedsOverlayStartupPrereqs: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['commandNeedsOverlayStartupPrereqs'];
    commandNeedsOverlayRuntime: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['commandNeedsOverlayRuntime'];
    handleCliCommandRuntimeServiceWithContext: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['handleCliCommandRuntimeServiceWithContext'];
    shouldStartApp: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['shouldStartApp'];
    parseArgs: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['parseArgs'];
    printHelp: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['printHelp'];
    onWillQuitCleanupHandler: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['onWillQuitCleanupHandler'];
    shouldRestoreWindowsOnActivateHandler: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['shouldRestoreWindowsOnActivateHandler'];
    restoreWindowsOnActivateHandler: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['restoreWindowsOnActivateHandler'];
    forceX11Backend: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['forceX11Backend'];
    enforceUnsupportedWaylandMode: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['enforceUnsupportedWaylandMode'];
    getDefaultSocketPath: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['getDefaultSocketPathHandler'];
    generateDefaultConfigFile: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['generateDefaultConfigFile'];
    runStartupBootstrapRuntime: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['runStartupBootstrapRuntime'];
    applyStartupState: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['applyStartupState'];
    getStartupModeFlags: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['getStartupModeFlags'];
    requestAppQuit: MainStartupRuntimeBootstrapInput<TStartupState>['commands']['requestAppQuit'];
  };
  constants: MainStartupRuntimeBootstrapInput<TStartupState>['constants'];
}

export function createMainStartupRuntimeBootstrap<TStartupState>(
  input: MainStartupRuntimeBootstrapInput<TStartupState>,
): MainStartupRuntimeBootstrap<TStartupState> {
  const startupRuntime = createMainStartupBootstrap<TStartupState>({
    appState: input.appState,
    appLifecycle: {
      app: input.appLifecycle.app,
      argv: input.appLifecycle.argv,
      platform: input.appLifecycle.platform,
    },
    config: input.config,
    logging: input.logging,
    shell: input.shell,
    runtime: {
      ...input.runtime,
      overlayUi: {
        get: () => input.runtime.getOverlayUi(),
      },
      yomitan: {
        loadYomitanExtension: () => input.runtime.getYomitanRuntime().loadYomitanExtension(),
        ensureYomitanExtensionLoaded: () =>
          input.runtime.getYomitanRuntime().ensureYomitanExtensionLoaded(),
        openYomitanSettings: () => input.runtime.getYomitanRuntime().openYomitanSettings(),
      },
    },
    commands: {
      ...input.commands,
      startTexthooker: (port, websocketUrl) => {
        if (!input.runtime.texthookerService.isRunning()) {
          input.runtime.texthookerService.start(port, websocketUrl);
        }
      },
      generateCharacterDictionary: async (targetPath?: string) => {
        const disabledReason = input.runtime.getCharacterDictionaryDisabledReason();
        if (disabledReason) {
          throw new Error(disabledReason);
        }
        return await input.runtime.dictionarySupport.generateCharacterDictionaryForCurrentMedia(
          targetPath,
        );
      },
      runYoutubePlaybackFlow: (request) => input.commands.runYoutubePlaybackFlow(request),
      getMultiCopyTimeoutMs: () => input.commands.getConfiguredShortcuts().multiCopyTimeoutMs,
    },
    constants: input.constants,
  });

  return {
    startupRuntime,
  };
}

export function createMainStartupRuntimeFromMainState<TStartupState>(
  input: MainStartupRuntimeFromMainStateInput<TStartupState>,
): MainStartupRuntimeBootstrap<TStartupState> {
  return createMainStartupRuntimeBootstrap<TStartupState>({
    appState: input.appState,
    appLifecycle: input.appLifecycle,
    config: input.config,
    logging: input.logging,
    shell: input.shell,
    runtime: {
      subtitle: input.runtime.subtitle,
      getOverlayUi: () => input.runtime.getOverlayUi(),
      overlayManager: input.runtime.overlayManager,
      firstRun: {
        ensureSetupStateInitialized: () => input.runtime.firstRun.ensureSetupStateInitialized(),
        openFirstRunSetupWindow: () => input.runtime.firstRun.openFirstRunSetupWindow(),
      },
      anilist: {
        refreshAnilistClientSecretStateIfEnabled: (options) =>
          input.runtime.anilist.refreshAnilistClientSecretStateIfEnabled(options),
        openAnilistSetupWindow: () => input.runtime.anilist.openAnilistSetupWindow(),
        getStatusSnapshot: () => input.runtime.anilist.getStatusSnapshot(),
        clearTokenState: () => input.runtime.anilist.clearTokenState(),
        getQueueStatusSnapshot: () => input.runtime.anilist.getQueueStatusSnapshot(),
        processNextAnilistRetryUpdate: () => input.runtime.anilist.processNextAnilistRetryUpdate(),
      },
      jellyfin: input.runtime.jellyfin,
      stats: input.runtime.stats,
      mining: input.runtime.mining,
      texthookerService: input.runtime.texthookerService,
      getYomitanRuntime: () => input.runtime.yomitan,
      getCharacterDictionaryDisabledReason: () =>
        input.runtime.getCharacterDictionaryDisabledReason(),
      subsyncRuntime: input.runtime.subsyncRuntime,
      dictionarySupport: input.runtime.dictionarySupport,
      subtitleWsService: input.runtime.subtitleWsService,
      annotationSubtitleWsService: input.runtime.annotationSubtitleWsService,
      immersion: input.runtime.immersion,
    },
    commands: {
      createMpvClientRuntimeService: () =>
        input.commands.mpvRuntime.createMpvClientRuntimeService(),
      createMecabTokenizerAndCheck: () => input.commands.mpvRuntime.createMecabTokenizerAndCheck(),
      prewarmSubtitleDictionaries: () => input.commands.mpvRuntime.prewarmSubtitleDictionaries(),
      startBackgroundWarmupsIfAllowed: () => input.commands.mpvRuntime.startBackgroundWarmups(),
      startBackgroundWarmups: () => input.commands.mpvRuntime.startBackgroundWarmups(),
      runHeadlessInitialCommand: () => input.commands.runHeadlessInitialCommand(),
      startPendingMultiCopy: (timeoutMs) =>
        input.commands.shortcuts.startPendingMultiCopy(timeoutMs),
      startPendingMineSentenceMultiple: (timeoutMs) =>
        input.commands.shortcuts.startPendingMineSentenceMultiple(timeoutMs),
      cycleSecondarySubMode: () => input.commands.cycleSecondarySubMode(),
      refreshOverlayShortcuts: () => input.commands.shortcuts.refreshOverlayShortcuts(),
      hasMpvWebsocketPlugin: () => input.commands.hasMpvWebsocketPlugin(),
      showMpvOsd: (text) => input.commands.showMpvOsd(text),
      shouldAutoOpenFirstRunSetup: (args) => input.commands.shouldAutoOpenFirstRunSetup(args),
      getConfiguredShortcuts: () => input.commands.shortcuts.getConfiguredShortcuts(),
      runYoutubePlaybackFlow: (request) => input.commands.youtube.runYoutubePlaybackFlow(request),
      shouldEnsureTrayOnStartupForInitialArgs: (platform, initialArgs) =>
        input.commands.shouldEnsureTrayOnStartupForInitialArgs(platform, initialArgs),
      isHeadlessInitialCommand: (args) => input.commands.isHeadlessInitialCommand(args),
      commandNeedsOverlayStartupPrereqs: (args) =>
        input.commands.commandNeedsOverlayStartupPrereqs(args),
      commandNeedsOverlayRuntime: (args) => input.commands.commandNeedsOverlayRuntime(args),
      handleCliCommandRuntimeServiceWithContext: (args, source, cliContext) =>
        input.commands.handleCliCommandRuntimeServiceWithContext(args, source, cliContext),
      shouldStartApp: (args) => input.commands.shouldStartApp(args),
      parseArgs: (argv) => input.commands.parseArgs(argv),
      printHelp: input.commands.printHelp,
      onWillQuitCleanupHandler: () => input.commands.onWillQuitCleanupHandler(),
      shouldRestoreWindowsOnActivateHandler: () =>
        input.commands.shouldRestoreWindowsOnActivateHandler(),
      restoreWindowsOnActivateHandler: () => input.commands.restoreWindowsOnActivateHandler(),
      forceX11Backend: (args) => input.commands.forceX11Backend(args),
      enforceUnsupportedWaylandMode: (args) => input.commands.enforceUnsupportedWaylandMode(args),
      getDefaultSocketPathHandler: () => input.commands.getDefaultSocketPath(),
      generateDefaultConfigFile: input.commands.generateDefaultConfigFile,
      runStartupBootstrapRuntime: (deps) => input.commands.runStartupBootstrapRuntime(deps),
      applyStartupState: (startupState) => input.commands.applyStartupState(startupState),
      getStartupModeFlags: input.commands.getStartupModeFlags,
      requestAppQuit: input.commands.requestAppQuit,
    },
    constants: input.constants,
  });
}

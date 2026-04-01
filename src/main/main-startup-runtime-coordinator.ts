import type { AnilistRuntime } from './anilist-runtime';
import type { DictionarySupportRuntime } from './dictionary-support-runtime';
import type { FirstRunRuntime } from './first-run-runtime';
import type { JellyfinRuntime } from './jellyfin-runtime';
import {
  createMainStartupRuntimeFromMainState,
  type MainStartupRuntimeBootstrap,
  type MainStartupRuntimeFromMainStateInput,
} from './main-startup-runtime-bootstrap';
import type { MiningRuntime } from './mining-runtime';
import type { MpvRuntime } from './mpv-runtime';
import type { ShortcutsRuntime } from './shortcuts-runtime';
import type { SubtitleRuntime } from './subtitle-runtime';
import type { YoutubeRuntime } from './youtube-runtime';

export interface MainStartupRuntimeCoordinatorInput<TStartupState> {
  appState: MainStartupRuntimeFromMainStateInput<TStartupState>['appState'];
  appLifecycle: MainStartupRuntimeFromMainStateInput<TStartupState>['appLifecycle'];
  config: MainStartupRuntimeFromMainStateInput<TStartupState>['config'];
  logging: MainStartupRuntimeFromMainStateInput<TStartupState>['logging'];
  shell: MainStartupRuntimeFromMainStateInput<TStartupState>['shell'];
  runtime: {
    subtitle: SubtitleRuntime;
    getOverlayUi: MainStartupRuntimeFromMainStateInput<TStartupState>['runtime']['getOverlayUi'];
    overlayManager: MainStartupRuntimeFromMainStateInput<TStartupState>['runtime']['overlayManager'];
    firstRun: Pick<FirstRunRuntime, 'ensureSetupStateInitialized' | 'openFirstRunSetupWindow'>;
    anilist: AnilistRuntime;
    jellyfin: JellyfinRuntime;
    stats: {
      ensureImmersionTrackerStarted: MainStartupRuntimeFromMainStateInput<TStartupState>['runtime']['stats']['ensureImmersionTrackerStarted'];
      runStatsCliCommand: MainStartupRuntimeFromMainStateInput<TStartupState>['runtime']['stats']['runStatsCliCommand'];
      immersion: MainStartupRuntimeFromMainStateInput<TStartupState>['runtime']['immersion'];
    };
    mining: {
      copyCurrentSubtitle: Pick<MiningRuntime, 'copyCurrentSubtitle'>['copyCurrentSubtitle'];
      markLastCardAsAudioCard: Pick<
        MiningRuntime,
        'markLastCardAsAudioCard'
      >['markLastCardAsAudioCard'];
      mineSentenceCard: Pick<MiningRuntime, 'mineSentenceCard'>['mineSentenceCard'];
      refreshKnownWordCache: Pick<MiningRuntime, 'refreshKnownWordCache'>['refreshKnownWordCache'];
      triggerFieldGrouping: Pick<MiningRuntime, 'triggerFieldGrouping'>['triggerFieldGrouping'];
      updateLastCardFromClipboard: Pick<
        MiningRuntime,
        'updateLastCardFromClipboard'
      >['updateLastCardFromClipboard'];
    };
    texthookerService: MainStartupRuntimeFromMainStateInput<TStartupState>['runtime']['texthookerService'];
    yomitan: {
      loadYomitanExtension: () => Promise<unknown>;
      ensureYomitanExtensionLoaded: () => Promise<unknown>;
      openYomitanSettings: () => boolean;
      getCharacterDictionaryDisabledReason: () => string | null;
    };
    subsyncRuntime: MainStartupRuntimeFromMainStateInput<TStartupState>['runtime']['subsyncRuntime'];
    dictionarySupport: DictionarySupportRuntime;
    subtitleWsService: MainStartupRuntimeFromMainStateInput<TStartupState>['runtime']['subtitleWsService'];
    annotationSubtitleWsService: MainStartupRuntimeFromMainStateInput<TStartupState>['runtime']['annotationSubtitleWsService'];
  };
  commands: {
    mpvRuntime: {
      createMpvClientRuntimeService: Pick<
        MpvRuntime,
        'createMpvClientRuntimeService'
      >['createMpvClientRuntimeService'];
      createMecabTokenizerAndCheck: Pick<
        MpvRuntime,
        'createMecabTokenizerAndCheck'
      >['createMecabTokenizerAndCheck'];
      prewarmSubtitleDictionaries: Pick<
        MpvRuntime,
        'prewarmSubtitleDictionaries'
      >['prewarmSubtitleDictionaries'];
      startBackgroundWarmups: Pick<MpvRuntime, 'startBackgroundWarmups'>['startBackgroundWarmups'];
    };
    runHeadlessInitialCommand: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['runHeadlessInitialCommand'];
    shortcuts: Pick<
      ShortcutsRuntime,
      | 'getConfiguredShortcuts'
      | 'refreshOverlayShortcuts'
      | 'startPendingMineSentenceMultiple'
      | 'startPendingMultiCopy'
    >;
    cycleSecondarySubMode: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['cycleSecondarySubMode'];
    hasMpvWebsocketPlugin: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['hasMpvWebsocketPlugin'];
    showMpvOsd: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['showMpvOsd'];
    shouldAutoOpenFirstRunSetup: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['shouldAutoOpenFirstRunSetup'];
    youtube: Pick<YoutubeRuntime, 'runYoutubePlaybackFlow'>;
    shouldEnsureTrayOnStartupForInitialArgs: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['shouldEnsureTrayOnStartupForInitialArgs'];
    isHeadlessInitialCommand: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['isHeadlessInitialCommand'];
    commandNeedsOverlayStartupPrereqs: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['commandNeedsOverlayStartupPrereqs'];
    commandNeedsOverlayRuntime: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['commandNeedsOverlayRuntime'];
    handleCliCommandRuntimeServiceWithContext: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['handleCliCommandRuntimeServiceWithContext'];
    shouldStartApp: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['shouldStartApp'];
    parseArgs: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['parseArgs'];
    printHelp: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['printHelp'];
    onWillQuitCleanupHandler: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['onWillQuitCleanupHandler'];
    shouldRestoreWindowsOnActivateHandler: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['shouldRestoreWindowsOnActivateHandler'];
    restoreWindowsOnActivateHandler: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['restoreWindowsOnActivateHandler'];
    forceX11Backend: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['forceX11Backend'];
    enforceUnsupportedWaylandMode: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['enforceUnsupportedWaylandMode'];
    getDefaultSocketPath: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['getDefaultSocketPath'];
    generateDefaultConfigFile: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['generateDefaultConfigFile'];
    runStartupBootstrapRuntime: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['runStartupBootstrapRuntime'];
    applyStartupState: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['applyStartupState'];
    getStartupModeFlags: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['getStartupModeFlags'];
    requestAppQuit: MainStartupRuntimeFromMainStateInput<TStartupState>['commands']['requestAppQuit'];
  };
  constants: MainStartupRuntimeFromMainStateInput<TStartupState>['constants'];
}

export interface MainStartupRuntimeFromProcessStateInput<TStartupState> {
  appState: MainStartupRuntimeCoordinatorInput<TStartupState>['appState'];
  appLifecycle: MainStartupRuntimeCoordinatorInput<TStartupState>['appLifecycle'];
  config: MainStartupRuntimeCoordinatorInput<TStartupState>['config'];
  logging: MainStartupRuntimeCoordinatorInput<TStartupState>['logging'];
  shell: MainStartupRuntimeCoordinatorInput<TStartupState>['shell'];
  runtime: {
    subtitle: SubtitleRuntime;
    startupOverlayUiAdapter: MainStartupRuntimeCoordinatorInput<TStartupState>['runtime']['getOverlayUi'] extends () => infer T
      ? T
      : never;
    overlayManager: MainStartupRuntimeCoordinatorInput<TStartupState>['runtime']['overlayManager'];
    firstRun: Pick<FirstRunRuntime, 'ensureSetupStateInitialized' | 'openFirstRunSetupWindow'>;
    anilist: AnilistRuntime;
    jellyfin: JellyfinRuntime;
    stats: {
      ensureImmersionTrackerStarted: MainStartupRuntimeCoordinatorInput<TStartupState>['runtime']['stats']['ensureImmersionTrackerStarted'];
      runStatsCliCommand: MainStartupRuntimeCoordinatorInput<TStartupState>['runtime']['stats']['runStatsCliCommand'];
      immersion: MainStartupRuntimeCoordinatorInput<TStartupState>['runtime']['stats']['immersion'];
    };
    mining: MiningRuntime;
    texthookerService: MainStartupRuntimeCoordinatorInput<TStartupState>['runtime']['texthookerService'];
    yomitan: {
      loadYomitanExtension: () => Promise<unknown>;
      ensureYomitanExtensionLoaded: () => Promise<unknown>;
      openYomitanSettings: () => boolean;
      getCharacterDictionaryDisabledReason: () => string | null;
    };
    subsyncRuntime: MainStartupRuntimeCoordinatorInput<TStartupState>['runtime']['subsyncRuntime'];
    dictionarySupport: DictionarySupportRuntime;
    subtitleWsService: MainStartupRuntimeCoordinatorInput<TStartupState>['runtime']['subtitleWsService'];
    annotationSubtitleWsService: MainStartupRuntimeCoordinatorInput<TStartupState>['runtime']['annotationSubtitleWsService'];
  };
  commands: {
    mpvRuntime: MpvRuntime;
    runHeadlessInitialCommand: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['runHeadlessInitialCommand'];
    shortcuts: Pick<
      ShortcutsRuntime,
      | 'getConfiguredShortcuts'
      | 'refreshOverlayShortcuts'
      | 'startPendingMineSentenceMultiple'
      | 'startPendingMultiCopy'
    >;
    cycleSecondarySubMode: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['cycleSecondarySubMode'];
    hasMpvWebsocketPlugin: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['hasMpvWebsocketPlugin'];
    showMpvOsd: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['showMpvOsd'];
    shouldAutoOpenFirstRunSetup: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['shouldAutoOpenFirstRunSetup'];
    youtube: YoutubeRuntime;
    shouldEnsureTrayOnStartupForInitialArgs: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['shouldEnsureTrayOnStartupForInitialArgs'];
    isHeadlessInitialCommand: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['isHeadlessInitialCommand'];
    commandNeedsOverlayStartupPrereqs: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['commandNeedsOverlayStartupPrereqs'];
    commandNeedsOverlayRuntime: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['commandNeedsOverlayRuntime'];
    handleCliCommandRuntimeServiceWithContext: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['handleCliCommandRuntimeServiceWithContext'];
    shouldStartApp: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['shouldStartApp'];
    parseArgs: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['parseArgs'];
    printHelp: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['printHelp'];
    onWillQuitCleanupHandler: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['onWillQuitCleanupHandler'];
    shouldRestoreWindowsOnActivateHandler: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['shouldRestoreWindowsOnActivateHandler'];
    restoreWindowsOnActivateHandler: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['restoreWindowsOnActivateHandler'];
    forceX11Backend: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['forceX11Backend'];
    enforceUnsupportedWaylandMode: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['enforceUnsupportedWaylandMode'];
    getDefaultSocketPath: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['getDefaultSocketPath'];
    generateDefaultConfigFile: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['generateDefaultConfigFile'];
    runStartupBootstrapRuntime: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['runStartupBootstrapRuntime'];
    applyStartupState: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['applyStartupState'];
    getStartupModeFlags: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['getStartupModeFlags'];
    requestAppQuit: MainStartupRuntimeCoordinatorInput<TStartupState>['commands']['requestAppQuit'];
  };
  constants: MainStartupRuntimeCoordinatorInput<TStartupState>['constants'];
}

export function createMainStartupRuntimeCoordinator<TStartupState>(
  input: MainStartupRuntimeCoordinatorInput<TStartupState>,
): MainStartupRuntimeBootstrap<TStartupState> {
  return createMainStartupRuntimeFromMainState<TStartupState>({
    appState: input.appState,
    appLifecycle: input.appLifecycle,
    config: input.config,
    logging: input.logging,
    shell: input.shell,
    runtime: {
      subtitle: {
        loadSubtitlePosition: () => input.runtime.subtitle.loadSubtitlePosition(),
        invalidateTokenizationCache: () => {
          input.runtime.subtitle.invalidateTokenizationCache();
        },
        refreshSubtitlePrefetchFromActiveTrack: () =>
          input.runtime.subtitle.refreshSubtitlePrefetchFromActiveTrack(),
      },
      getOverlayUi: () => input.runtime.getOverlayUi(),
      overlayManager: input.runtime.overlayManager,
      firstRun: input.runtime.firstRun,
      anilist: input.runtime.anilist,
      jellyfin: {
        startJellyfinRemoteSession: () => input.runtime.jellyfin.startJellyfinRemoteSession(),
        openJellyfinSetupWindow: () => input.runtime.jellyfin.openJellyfinSetupWindow(),
        runJellyfinCommand: (argsFromCommand) =>
          input.runtime.jellyfin.runJellyfinCommand(argsFromCommand),
      },
      stats: {
        ensureImmersionTrackerStarted: () => input.runtime.stats.ensureImmersionTrackerStarted(),
        runStatsCliCommand: (argsFromCommand, source) =>
          input.runtime.stats.runStatsCliCommand(argsFromCommand, source),
      },
      mining: {
        copyCurrentSubtitle: () => input.runtime.mining.copyCurrentSubtitle(),
        mineSentenceCard: () => input.runtime.mining.mineSentenceCard(),
        updateLastCardFromClipboard: () => input.runtime.mining.updateLastCardFromClipboard(),
        refreshKnownWordCache: () => input.runtime.mining.refreshKnownWordCache(),
        triggerFieldGrouping: () => input.runtime.mining.triggerFieldGrouping(),
        markLastCardAsAudioCard: () => input.runtime.mining.markLastCardAsAudioCard(),
      },
      texthookerService: input.runtime.texthookerService,
      yomitan: {
        loadYomitanExtension: () => input.runtime.yomitan.loadYomitanExtension(),
        ensureYomitanExtensionLoaded: () => input.runtime.yomitan.ensureYomitanExtensionLoaded(),
        openYomitanSettings: () => input.runtime.yomitan.openYomitanSettings(),
      },
      getCharacterDictionaryDisabledReason: () =>
        input.runtime.yomitan.getCharacterDictionaryDisabledReason(),
      subsyncRuntime: input.runtime.subsyncRuntime,
      dictionarySupport: input.runtime.dictionarySupport,
      subtitleWsService: input.runtime.subtitleWsService,
      annotationSubtitleWsService: input.runtime.annotationSubtitleWsService,
      immersion: input.runtime.stats.immersion,
    },
    commands: {
      mpvRuntime: {
        createMpvClientRuntimeService: () =>
          input.commands.mpvRuntime.createMpvClientRuntimeService(),
        createMecabTokenizerAndCheck: () =>
          input.commands.mpvRuntime.createMecabTokenizerAndCheck(),
        prewarmSubtitleDictionaries: () => input.commands.mpvRuntime.prewarmSubtitleDictionaries(),
        startBackgroundWarmups: () => input.commands.mpvRuntime.startBackgroundWarmups(),
      },
      runHeadlessInitialCommand: () => input.commands.runHeadlessInitialCommand(),
      shortcuts: {
        startPendingMultiCopy: (timeoutMs) =>
          input.commands.shortcuts.startPendingMultiCopy(timeoutMs),
        startPendingMineSentenceMultiple: (timeoutMs) =>
          input.commands.shortcuts.startPendingMineSentenceMultiple(timeoutMs),
        refreshOverlayShortcuts: () => input.commands.shortcuts.refreshOverlayShortcuts(),
        getConfiguredShortcuts: () => input.commands.shortcuts.getConfiguredShortcuts(),
      },
      cycleSecondarySubMode: () => input.commands.cycleSecondarySubMode(),
      hasMpvWebsocketPlugin: () => input.commands.hasMpvWebsocketPlugin(),
      showMpvOsd: (text) => input.commands.showMpvOsd(text),
      shouldAutoOpenFirstRunSetup: (args) => input.commands.shouldAutoOpenFirstRunSetup(args),
      youtube: {
        runYoutubePlaybackFlow: (request) => input.commands.youtube.runYoutubePlaybackFlow(request),
      },
      shouldEnsureTrayOnStartupForInitialArgs: (platform, initialArgs) =>
        input.commands.shouldEnsureTrayOnStartupForInitialArgs(platform, initialArgs ?? null),
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
      getDefaultSocketPath: () => input.commands.getDefaultSocketPath(),
      generateDefaultConfigFile: input.commands.generateDefaultConfigFile,
      runStartupBootstrapRuntime: (deps) => input.commands.runStartupBootstrapRuntime(deps),
      applyStartupState: (startupState) => input.commands.applyStartupState(startupState),
      getStartupModeFlags: input.commands.getStartupModeFlags,
      requestAppQuit: input.commands.requestAppQuit,
    },
    constants: input.constants,
  });
}

export function createMainStartupRuntimeFromProcessState<TStartupState>(
  input: MainStartupRuntimeFromProcessStateInput<TStartupState>,
) {
  return createMainStartupRuntimeCoordinator<TStartupState>({
    appState: input.appState,
    appLifecycle: input.appLifecycle,
    config: input.config,
    logging: input.logging,
    shell: input.shell,
    runtime: {
      subtitle: input.runtime.subtitle,
      getOverlayUi: () => input.runtime.startupOverlayUiAdapter,
      overlayManager: input.runtime.overlayManager,
      firstRun: input.runtime.firstRun,
      anilist: input.runtime.anilist,
      jellyfin: input.runtime.jellyfin,
      stats: {
        ensureImmersionTrackerStarted: () => input.runtime.stats.ensureImmersionTrackerStarted(),
        runStatsCliCommand: (argsFromCommand, source) =>
          input.runtime.stats.runStatsCliCommand(argsFromCommand, source),
        immersion: input.runtime.stats.immersion,
      },
      mining: {
        copyCurrentSubtitle: () => input.runtime.mining.copyCurrentSubtitle(),
        markLastCardAsAudioCard: () => input.runtime.mining.markLastCardAsAudioCard(),
        mineSentenceCard: () => input.runtime.mining.mineSentenceCard(),
        refreshKnownWordCache: () => input.runtime.mining.refreshKnownWordCache(),
        triggerFieldGrouping: () => input.runtime.mining.triggerFieldGrouping(),
        updateLastCardFromClipboard: () => input.runtime.mining.updateLastCardFromClipboard(),
      },
      texthookerService: input.runtime.texthookerService,
      yomitan: input.runtime.yomitan,
      subsyncRuntime: input.runtime.subsyncRuntime,
      dictionarySupport: input.runtime.dictionarySupport,
      subtitleWsService: input.runtime.subtitleWsService,
      annotationSubtitleWsService: input.runtime.annotationSubtitleWsService,
    },
    commands: {
      mpvRuntime: {
        createMpvClientRuntimeService: () =>
          input.commands.mpvRuntime.createMpvClientRuntimeService(),
        createMecabTokenizerAndCheck: () =>
          input.commands.mpvRuntime.createMecabTokenizerAndCheck(),
        prewarmSubtitleDictionaries: () => input.commands.mpvRuntime.prewarmSubtitleDictionaries(),
        startBackgroundWarmups: () => input.commands.mpvRuntime.startBackgroundWarmups(),
      },
      runHeadlessInitialCommand: () => input.commands.runHeadlessInitialCommand(),
      shortcuts: input.commands.shortcuts,
      cycleSecondarySubMode: () => input.commands.cycleSecondarySubMode(),
      hasMpvWebsocketPlugin: () => input.commands.hasMpvWebsocketPlugin(),
      showMpvOsd: (text) => input.commands.showMpvOsd(text),
      shouldAutoOpenFirstRunSetup: (args) => input.commands.shouldAutoOpenFirstRunSetup(args),
      youtube: input.commands.youtube,
      shouldEnsureTrayOnStartupForInitialArgs: (platform, initialArgs) =>
        input.commands.shouldEnsureTrayOnStartupForInitialArgs(platform, initialArgs ?? null),
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
      getDefaultSocketPath: () => input.commands.getDefaultSocketPath(),
      generateDefaultConfigFile: input.commands.generateDefaultConfigFile,
      runStartupBootstrapRuntime: (deps) => input.commands.runStartupBootstrapRuntime(deps),
      applyStartupState: (startupState) => input.commands.applyStartupState(startupState),
      getStartupModeFlags: input.commands.getStartupModeFlags,
      requestAppQuit: input.commands.requestAppQuit,
    },
    constants: input.constants,
  }).startupRuntime;
}

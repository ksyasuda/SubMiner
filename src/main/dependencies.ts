import { RuntimeOptionId, RuntimeOptionValue, SubsyncManualPayload } from '../types';
import type { OverlayNotificationPayload } from '../types/notification';
import { SubsyncResolvedConfig } from '../subsync/utils';
import type { SubsyncRuntimeDeps } from '../core/services/subsync-runner';
import type { IpcDepsRuntimeOptions } from '../core/services/ipc';
import type { AnkiJimakuIpcRuntimeOptions } from '../core/services/anki-jimaku';
import type { CliCommandDepsRuntimeOptions } from '../core/services/cli-command';
import type { HandleMpvCommandFromIpcOptions } from '../core/services/ipc-command';
import {
  cycleRuntimeOptionFromIpcRuntime,
  setRuntimeOptionFromIpcRuntime,
} from '../core/services/runtime-options-ipc';
import { RuntimeOptionsManager } from '../runtime-options';

export interface RuntimeOptionsIpcDepsParams {
  getRuntimeOptionsManager: () => RuntimeOptionsManager | null;
  showMpvOsd: (text: string) => void;
}

export interface SubsyncRuntimeDepsParams {
  getMpvClient: () => ReturnType<SubsyncRuntimeDeps['getMpvClient']>;
  getResolvedSubsyncConfig: () => SubsyncResolvedConfig;
  isSubsyncInProgress: () => boolean;
  setSubsyncInProgress: (inProgress: boolean) => void;
  showMpvOsd: (text: string) => void;
  openManualPicker: (payload: SubsyncManualPayload) => void;
}

export function createRuntimeOptionsIpcDeps(params: RuntimeOptionsIpcDepsParams): {
  setRuntimeOption: (id: RuntimeOptionId, value: RuntimeOptionValue) => unknown;
  cycleRuntimeOption: (id: RuntimeOptionId, direction: 1 | -1) => unknown;
} {
  return {
    setRuntimeOption: (id, value) =>
      setRuntimeOptionFromIpcRuntime(params.getRuntimeOptionsManager(), id, value, (text) =>
        params.showMpvOsd(text),
      ),
    cycleRuntimeOption: (id, direction) =>
      cycleRuntimeOptionFromIpcRuntime(params.getRuntimeOptionsManager(), id, direction, (text) =>
        params.showMpvOsd(text),
      ),
  };
}

export function createSubsyncRuntimeDeps(params: SubsyncRuntimeDepsParams): SubsyncRuntimeDeps {
  return {
    getMpvClient: params.getMpvClient,
    getResolvedSubsyncConfig: params.getResolvedSubsyncConfig,
    isSubsyncInProgress: params.isSubsyncInProgress,
    setSubsyncInProgress: params.setSubsyncInProgress,
    showMpvOsd: params.showMpvOsd,
    openManualPicker: params.openManualPicker,
  };
}

export interface MainIpcRuntimeServiceDepsParams {
  getMainWindow: IpcDepsRuntimeOptions['getMainWindow'];
  getVisibleOverlayVisibility: IpcDepsRuntimeOptions['getVisibleOverlayVisibility'];
  onOverlayModalClosed: IpcDepsRuntimeOptions['onOverlayModalClosed'];
  onOverlayModalOpened?: IpcDepsRuntimeOptions['onOverlayModalOpened'];
  onOverlayMouseInteractionChanged?: IpcDepsRuntimeOptions['onOverlayMouseInteractionChanged'];
  onOverlayInteractiveHint?: IpcDepsRuntimeOptions['onOverlayInteractiveHint'];
  handleOverlayNotificationAction?: IpcDepsRuntimeOptions['handleOverlayNotificationAction'];
  onYoutubePickerResolve: IpcDepsRuntimeOptions['onYoutubePickerResolve'];
  openYomitanSettings: IpcDepsRuntimeOptions['openYomitanSettings'];
  quitApp: IpcDepsRuntimeOptions['quitApp'];
  toggleVisibleOverlay: IpcDepsRuntimeOptions['toggleVisibleOverlay'];
  tokenizeCurrentSubtitle: IpcDepsRuntimeOptions['tokenizeCurrentSubtitle'];
  getCurrentSubtitleRaw: IpcDepsRuntimeOptions['getCurrentSubtitleRaw'];
  getCurrentSubtitleAss: IpcDepsRuntimeOptions['getCurrentSubtitleAss'];
  getSubtitleSidebarSnapshot?: IpcDepsRuntimeOptions['getSubtitleSidebarSnapshot'];
  getSubtitleSidebarOpen?: IpcDepsRuntimeOptions['getSubtitleSidebarOpen'];
  getPlaybackPaused: IpcDepsRuntimeOptions['getPlaybackPaused'];
  focusMainWindow?: IpcDepsRuntimeOptions['focusMainWindow'];
  activatePlaybackWindowForOverlayInteraction?: IpcDepsRuntimeOptions['activatePlaybackWindowForOverlayInteraction'];
  getSubtitlePosition: IpcDepsRuntimeOptions['getSubtitlePosition'];
  getSubtitleStyle: IpcDepsRuntimeOptions['getSubtitleStyle'];
  saveSubtitlePosition: IpcDepsRuntimeOptions['saveSubtitlePosition'];
  getMecabTokenizer: IpcDepsRuntimeOptions['getMecabTokenizer'];
  handleMpvCommand: IpcDepsRuntimeOptions['handleMpvCommand'];
  getKeybindings: IpcDepsRuntimeOptions['getKeybindings'];
  getSessionBindings: IpcDepsRuntimeOptions['getSessionBindings'];
  getConfiguredShortcuts: IpcDepsRuntimeOptions['getConfiguredShortcuts'];
  dispatchSessionAction: IpcDepsRuntimeOptions['dispatchSessionAction'];
  getStatsToggleKey: IpcDepsRuntimeOptions['getStatsToggleKey'];
  getMarkWatchedKey: IpcDepsRuntimeOptions['getMarkWatchedKey'];
  getOverlayNotificationPosition: IpcDepsRuntimeOptions['getOverlayNotificationPosition'];
  getControllerConfig: IpcDepsRuntimeOptions['getControllerConfig'];
  saveControllerConfig: IpcDepsRuntimeOptions['saveControllerConfig'];
  saveControllerPreference: IpcDepsRuntimeOptions['saveControllerPreference'];
  getSecondarySubMode: IpcDepsRuntimeOptions['getSecondarySubMode'];
  getMpvClient: IpcDepsRuntimeOptions['getMpvClient'];
  runSubsyncManual: IpcDepsRuntimeOptions['runSubsyncManual'];
  getAnkiConnectStatus: IpcDepsRuntimeOptions['getAnkiConnectStatus'];
  getRuntimeOptions: IpcDepsRuntimeOptions['getRuntimeOptions'];
  setRuntimeOption: IpcDepsRuntimeOptions['setRuntimeOption'];
  cycleRuntimeOption: IpcDepsRuntimeOptions['cycleRuntimeOption'];
  reportOverlayContentBounds: IpcDepsRuntimeOptions['reportOverlayContentBounds'];
  getAnilistStatus: IpcDepsRuntimeOptions['getAnilistStatus'];
  clearAnilistToken: IpcDepsRuntimeOptions['clearAnilistToken'];
  openAnilistSetup: IpcDepsRuntimeOptions['openAnilistSetup'];
  getAnilistQueueStatus: IpcDepsRuntimeOptions['getAnilistQueueStatus'];
  retryAnilistQueueNow: IpcDepsRuntimeOptions['retryAnilistQueueNow'];
  runAnilistPostWatchUpdateOnManualMark?: IpcDepsRuntimeOptions['runAnilistPostWatchUpdateOnManualMark'];
  recordSubtitleMiningContext?: IpcDepsRuntimeOptions['recordSubtitleMiningContext'];
  getCharacterDictionarySelection?: IpcDepsRuntimeOptions['getCharacterDictionarySelection'];
  setCharacterDictionarySelection?: IpcDepsRuntimeOptions['setCharacterDictionarySelection'];
  getCharacterDictionaryManagerSnapshot?: IpcDepsRuntimeOptions['getCharacterDictionaryManagerSnapshot'];
  removeCharacterDictionaryManagedEntry?: IpcDepsRuntimeOptions['removeCharacterDictionaryManagedEntry'];
  moveCharacterDictionaryManagedEntry?: IpcDepsRuntimeOptions['moveCharacterDictionaryManagedEntry'];
  appendClipboardVideoToQueue: IpcDepsRuntimeOptions['appendClipboardVideoToQueue'];
  getPlaylistBrowserSnapshot: IpcDepsRuntimeOptions['getPlaylistBrowserSnapshot'];
  appendPlaylistBrowserFile: IpcDepsRuntimeOptions['appendPlaylistBrowserFile'];
  playPlaylistBrowserIndex: IpcDepsRuntimeOptions['playPlaylistBrowserIndex'];
  removePlaylistBrowserIndex: IpcDepsRuntimeOptions['removePlaylistBrowserIndex'];
  movePlaylistBrowserIndex: IpcDepsRuntimeOptions['movePlaylistBrowserIndex'];
  getImmersionTracker?: IpcDepsRuntimeOptions['getImmersionTracker'];
}

export interface AnkiJimakuIpcRuntimeServiceDepsParams {
  patchAnkiConnectEnabled: AnkiJimakuIpcRuntimeOptions['patchAnkiConnectEnabled'];
  getResolvedConfig: AnkiJimakuIpcRuntimeOptions['getResolvedConfig'];
  getRuntimeOptionsManager: AnkiJimakuIpcRuntimeOptions['getRuntimeOptionsManager'];
  getSubtitleTimingTracker: AnkiJimakuIpcRuntimeOptions['getSubtitleTimingTracker'];
  getMpvClient: AnkiJimakuIpcRuntimeOptions['getMpvClient'];
  getAnkiIntegration: AnkiJimakuIpcRuntimeOptions['getAnkiIntegration'];
  setAnkiIntegration: AnkiJimakuIpcRuntimeOptions['setAnkiIntegration'];
  getKnownWordCacheStatePath: AnkiJimakuIpcRuntimeOptions['getKnownWordCacheStatePath'];
  getCachedMediaPath?: AnkiJimakuIpcRuntimeOptions['getCachedMediaPath'];
  shouldRequireRemoteMediaCache?: AnkiJimakuIpcRuntimeOptions['shouldRequireRemoteMediaCache'];
  showDesktopNotification: AnkiJimakuIpcRuntimeOptions['showDesktopNotification'];
  showOverlayNotification?: (payload: OverlayNotificationPayload) => void;
  createFieldGroupingCallback: AnkiJimakuIpcRuntimeOptions['createFieldGroupingCallback'];
  broadcastRuntimeOptionsChanged: AnkiJimakuIpcRuntimeOptions['broadcastRuntimeOptionsChanged'];
  getFieldGroupingResolver: AnkiJimakuIpcRuntimeOptions['getFieldGroupingResolver'];
  setFieldGroupingResolver: AnkiJimakuIpcRuntimeOptions['setFieldGroupingResolver'];
  parseMediaInfo: AnkiJimakuIpcRuntimeOptions['parseMediaInfo'];
  getCurrentMediaPath: AnkiJimakuIpcRuntimeOptions['getCurrentMediaPath'];
  jimakuFetchJson: AnkiJimakuIpcRuntimeOptions['jimakuFetchJson'];
  getJimakuMaxEntryResults: AnkiJimakuIpcRuntimeOptions['getJimakuMaxEntryResults'];
  getJimakuLanguagePreference: AnkiJimakuIpcRuntimeOptions['getJimakuLanguagePreference'];
  resolveJimakuApiKey: AnkiJimakuIpcRuntimeOptions['resolveJimakuApiKey'];
  isRemoteMediaPath: AnkiJimakuIpcRuntimeOptions['isRemoteMediaPath'];
  downloadToFile: AnkiJimakuIpcRuntimeOptions['downloadToFile'];
}

export interface CliCommandRuntimeServiceDepsParams {
  setLogLevel?: CliCommandDepsRuntimeOptions['setLogLevel'];
  mpv: {
    getSocketPath: CliCommandDepsRuntimeOptions['mpv']['getSocketPath'];
    setSocketPath: CliCommandDepsRuntimeOptions['mpv']['setSocketPath'];
    getClient: CliCommandDepsRuntimeOptions['mpv']['getClient'];
    showOsd: CliCommandDepsRuntimeOptions['mpv']['showOsd'];
    showPlaybackFeedback?: CliCommandDepsRuntimeOptions['mpv']['showPlaybackFeedback'];
  };
  texthooker: {
    service: CliCommandDepsRuntimeOptions['texthooker']['service'];
    getPort: CliCommandDepsRuntimeOptions['texthooker']['getPort'];
    setPort: CliCommandDepsRuntimeOptions['texthooker']['setPort'];
    getWebsocketUrl: CliCommandDepsRuntimeOptions['texthooker']['getWebsocketUrl'];
    shouldOpenBrowser: CliCommandDepsRuntimeOptions['texthooker']['shouldOpenBrowser'];
    openInBrowser: CliCommandDepsRuntimeOptions['texthooker']['openInBrowser'];
  };
  overlay: {
    isInitialized: CliCommandDepsRuntimeOptions['overlay']['isInitialized'];
    initialize: CliCommandDepsRuntimeOptions['overlay']['initialize'];
    toggleVisible: CliCommandDepsRuntimeOptions['overlay']['toggleVisible'];
    togglePrimarySubtitleBar: CliCommandDepsRuntimeOptions['overlay']['togglePrimarySubtitleBar'];
    setVisible: CliCommandDepsRuntimeOptions['overlay']['setVisible'];
  };
  mining: {
    copyCurrentSubtitle: CliCommandDepsRuntimeOptions['mining']['copyCurrentSubtitle'];
    startPendingMultiCopy: CliCommandDepsRuntimeOptions['mining']['startPendingMultiCopy'];
    mineSentenceCard: CliCommandDepsRuntimeOptions['mining']['mineSentenceCard'];
    startPendingMineSentenceMultiple: CliCommandDepsRuntimeOptions['mining']['startPendingMineSentenceMultiple'];
    updateLastCardFromClipboard: CliCommandDepsRuntimeOptions['mining']['updateLastCardFromClipboard'];
    refreshKnownWords: CliCommandDepsRuntimeOptions['mining']['refreshKnownWords'];
    triggerFieldGrouping: CliCommandDepsRuntimeOptions['mining']['triggerFieldGrouping'];
    triggerSubsyncFromConfig: CliCommandDepsRuntimeOptions['mining']['triggerSubsyncFromConfig'];
    markLastCardAsAudioCard: CliCommandDepsRuntimeOptions['mining']['markLastCardAsAudioCard'];
  };
  anilist: {
    getStatus: CliCommandDepsRuntimeOptions['anilist']['getStatus'];
    clearToken: CliCommandDepsRuntimeOptions['anilist']['clearToken'];
    openSetup: CliCommandDepsRuntimeOptions['anilist']['openSetup'];
    getQueueStatus: CliCommandDepsRuntimeOptions['anilist']['getQueueStatus'];
    retryQueueNow: CliCommandDepsRuntimeOptions['anilist']['retryQueueNow'];
  };
  dictionary: {
    generate: CliCommandDepsRuntimeOptions['dictionary']['generate'];
    getSelection: CliCommandDepsRuntimeOptions['dictionary']['getSelection'];
    setSelection: CliCommandDepsRuntimeOptions['dictionary']['setSelection'];
  };
  jellyfin: {
    openSetup: CliCommandDepsRuntimeOptions['jellyfin']['openSetup'];
    runStatsCommand: CliCommandDepsRuntimeOptions['jellyfin']['runStatsCommand'];
    runCommand: CliCommandDepsRuntimeOptions['jellyfin']['runCommand'];
  };
  app: {
    stop: CliCommandDepsRuntimeOptions['app']['stop'];
    hasMainWindow: CliCommandDepsRuntimeOptions['app']['hasMainWindow'];
    runUpdateCommand: CliCommandDepsRuntimeOptions['app']['runUpdateCommand'];
    runEnsureLinuxRuntimePluginAssetsCommand: CliCommandDepsRuntimeOptions['app']['runEnsureLinuxRuntimePluginAssetsCommand'];
    runYoutubePlaybackFlow: CliCommandDepsRuntimeOptions['app']['runYoutubePlaybackFlow'];
  };
  dispatchSessionAction: CliCommandDepsRuntimeOptions['dispatchSessionAction'];
  ui: {
    openFirstRunSetup: CliCommandDepsRuntimeOptions['ui']['openFirstRunSetup'];
    openYomitanSettings: CliCommandDepsRuntimeOptions['ui']['openYomitanSettings'];
    openConfigSettingsWindow: CliCommandDepsRuntimeOptions['ui']['openConfigSettingsWindow'];
    cycleSecondarySubMode: CliCommandDepsRuntimeOptions['ui']['cycleSecondarySubMode'];
    openRuntimeOptionsPalette: CliCommandDepsRuntimeOptions['ui']['openRuntimeOptionsPalette'];
    printHelp: CliCommandDepsRuntimeOptions['ui']['printHelp'];
  };
  getMultiCopyTimeoutMs: CliCommandDepsRuntimeOptions['getMultiCopyTimeoutMs'];
  schedule: CliCommandDepsRuntimeOptions['schedule'];
  log: CliCommandDepsRuntimeOptions['log'];
  logDebug: CliCommandDepsRuntimeOptions['logDebug'];
  warn: CliCommandDepsRuntimeOptions['warn'];
  error: CliCommandDepsRuntimeOptions['error'];
}

export interface MpvCommandRuntimeServiceDepsParams {
  specialCommands: HandleMpvCommandFromIpcOptions['specialCommands'];
  runtimeOptionsCycle: HandleMpvCommandFromIpcOptions['runtimeOptionsCycle'];
  triggerSubsyncFromConfig: HandleMpvCommandFromIpcOptions['triggerSubsyncFromConfig'];
  openRuntimeOptionsPalette: HandleMpvCommandFromIpcOptions['openRuntimeOptionsPalette'];
  openJimaku: HandleMpvCommandFromIpcOptions['openJimaku'];
  openYoutubeTrackPicker: HandleMpvCommandFromIpcOptions['openYoutubeTrackPicker'];
  openPlaylistBrowser: HandleMpvCommandFromIpcOptions['openPlaylistBrowser'];
  showMpvOsd: HandleMpvCommandFromIpcOptions['showMpvOsd'];
  showRawMpvOsd?: HandleMpvCommandFromIpcOptions['showRawMpvOsd'];
  showPlaybackFeedback?: HandleMpvCommandFromIpcOptions['showPlaybackFeedback'];
  mpvReplaySubtitle: HandleMpvCommandFromIpcOptions['mpvReplaySubtitle'];
  mpvPlayNextSubtitle: HandleMpvCommandFromIpcOptions['mpvPlayNextSubtitle'];
  mpvSendCommand: HandleMpvCommandFromIpcOptions['mpvSendCommand'];
  resolveProxyCommandOsd?: HandleMpvCommandFromIpcOptions['resolveProxyCommandOsd'];
  isMpvConnected: HandleMpvCommandFromIpcOptions['isMpvConnected'];
  hasRuntimeOptionsManager: HandleMpvCommandFromIpcOptions['hasRuntimeOptionsManager'];
}

export function createMainIpcRuntimeServiceDeps(
  params: MainIpcRuntimeServiceDepsParams,
): IpcDepsRuntimeOptions {
  return {
    getMainWindow: params.getMainWindow,
    getVisibleOverlayVisibility: params.getVisibleOverlayVisibility,
    onOverlayModalClosed: params.onOverlayModalClosed,
    onOverlayModalOpened: params.onOverlayModalOpened,
    onOverlayMouseInteractionChanged: params.onOverlayMouseInteractionChanged,
    onOverlayInteractiveHint: params.onOverlayInteractiveHint,
    handleOverlayNotificationAction: params.handleOverlayNotificationAction,
    onYoutubePickerResolve: params.onYoutubePickerResolve,
    openYomitanSettings: params.openYomitanSettings,
    quitApp: params.quitApp,
    toggleVisibleOverlay: params.toggleVisibleOverlay,
    tokenizeCurrentSubtitle: params.tokenizeCurrentSubtitle,
    getCurrentSubtitleRaw: params.getCurrentSubtitleRaw,
    getCurrentSubtitleAss: params.getCurrentSubtitleAss,
    getSubtitleSidebarSnapshot: params.getSubtitleSidebarSnapshot,
    getSubtitleSidebarOpen: params.getSubtitleSidebarOpen,
    getPlaybackPaused: params.getPlaybackPaused,
    getSubtitlePosition: params.getSubtitlePosition,
    getSubtitleStyle: params.getSubtitleStyle,
    saveSubtitlePosition: params.saveSubtitlePosition,
    getMecabTokenizer: params.getMecabTokenizer,
    handleMpvCommand: params.handleMpvCommand,
    getKeybindings: params.getKeybindings,
    getSessionBindings: params.getSessionBindings,
    getConfiguredShortcuts: params.getConfiguredShortcuts,
    dispatchSessionAction: params.dispatchSessionAction,
    getStatsToggleKey: params.getStatsToggleKey,
    getMarkWatchedKey: params.getMarkWatchedKey,
    getOverlayNotificationPosition: params.getOverlayNotificationPosition,
    getControllerConfig: params.getControllerConfig,
    saveControllerConfig: params.saveControllerConfig,
    saveControllerPreference: params.saveControllerPreference,
    focusMainWindow: params.focusMainWindow ?? (() => {}),
    activatePlaybackWindowForOverlayInteraction: params.activatePlaybackWindowForOverlayInteraction,
    getSecondarySubMode: params.getSecondarySubMode,
    getMpvClient: params.getMpvClient,
    runSubsyncManual: params.runSubsyncManual,
    getAnkiConnectStatus: params.getAnkiConnectStatus,
    getRuntimeOptions: params.getRuntimeOptions,
    setRuntimeOption: params.setRuntimeOption,
    cycleRuntimeOption: params.cycleRuntimeOption,
    reportOverlayContentBounds: params.reportOverlayContentBounds,
    getAnilistStatus: params.getAnilistStatus,
    clearAnilistToken: params.clearAnilistToken,
    openAnilistSetup: params.openAnilistSetup,
    getAnilistQueueStatus: params.getAnilistQueueStatus,
    retryAnilistQueueNow: params.retryAnilistQueueNow,
    runAnilistPostWatchUpdateOnManualMark: params.runAnilistPostWatchUpdateOnManualMark,
    recordSubtitleMiningContext: params.recordSubtitleMiningContext,
    getCharacterDictionarySelection: params.getCharacterDictionarySelection,
    setCharacterDictionarySelection: params.setCharacterDictionarySelection,
    getCharacterDictionaryManagerSnapshot: params.getCharacterDictionaryManagerSnapshot,
    removeCharacterDictionaryManagedEntry: params.removeCharacterDictionaryManagedEntry,
    moveCharacterDictionaryManagedEntry: params.moveCharacterDictionaryManagedEntry,
    appendClipboardVideoToQueue: params.appendClipboardVideoToQueue,
    getPlaylistBrowserSnapshot: params.getPlaylistBrowserSnapshot,
    appendPlaylistBrowserFile: params.appendPlaylistBrowserFile,
    playPlaylistBrowserIndex: params.playPlaylistBrowserIndex,
    removePlaylistBrowserIndex: params.removePlaylistBrowserIndex,
    movePlaylistBrowserIndex: params.movePlaylistBrowserIndex,
    getImmersionTracker: params.getImmersionTracker,
  };
}

export function createAnkiJimakuIpcRuntimeServiceDeps(
  params: AnkiJimakuIpcRuntimeServiceDepsParams,
): AnkiJimakuIpcRuntimeOptions {
  return {
    patchAnkiConnectEnabled: params.patchAnkiConnectEnabled,
    getResolvedConfig: params.getResolvedConfig,
    getRuntimeOptionsManager: params.getRuntimeOptionsManager,
    getSubtitleTimingTracker: params.getSubtitleTimingTracker,
    getMpvClient: params.getMpvClient,
    getAnkiIntegration: params.getAnkiIntegration,
    setAnkiIntegration: params.setAnkiIntegration,
    getKnownWordCacheStatePath: params.getKnownWordCacheStatePath,
    ...(params.getCachedMediaPath ? { getCachedMediaPath: params.getCachedMediaPath } : {}),
    ...(params.shouldRequireRemoteMediaCache
      ? { shouldRequireRemoteMediaCache: params.shouldRequireRemoteMediaCache }
      : {}),
    showDesktopNotification: params.showDesktopNotification,
    showOverlayNotification: params.showOverlayNotification,
    createFieldGroupingCallback: params.createFieldGroupingCallback,
    broadcastRuntimeOptionsChanged: params.broadcastRuntimeOptionsChanged,
    getFieldGroupingResolver: params.getFieldGroupingResolver,
    setFieldGroupingResolver: params.setFieldGroupingResolver,
    parseMediaInfo: params.parseMediaInfo,
    getCurrentMediaPath: params.getCurrentMediaPath,
    jimakuFetchJson: params.jimakuFetchJson,
    getJimakuMaxEntryResults: params.getJimakuMaxEntryResults,
    getJimakuLanguagePreference: params.getJimakuLanguagePreference,
    resolveJimakuApiKey: params.resolveJimakuApiKey,
    isRemoteMediaPath: params.isRemoteMediaPath,
    downloadToFile: params.downloadToFile,
  };
}

export function createCliCommandRuntimeServiceDeps(
  params: CliCommandRuntimeServiceDepsParams,
): CliCommandDepsRuntimeOptions {
  return {
    setLogLevel: params.setLogLevel,
    mpv: {
      getSocketPath: params.mpv.getSocketPath,
      setSocketPath: params.mpv.setSocketPath,
      getClient: params.mpv.getClient,
      showOsd: params.mpv.showOsd,
      showPlaybackFeedback: params.mpv.showPlaybackFeedback,
    },
    texthooker: {
      service: params.texthooker.service,
      getPort: params.texthooker.getPort,
      setPort: params.texthooker.setPort,
      getWebsocketUrl: params.texthooker.getWebsocketUrl,
      shouldOpenBrowser: params.texthooker.shouldOpenBrowser,
      openInBrowser: params.texthooker.openInBrowser,
    },
    overlay: {
      isInitialized: params.overlay.isInitialized,
      initialize: params.overlay.initialize,
      toggleVisible: params.overlay.toggleVisible,
      togglePrimarySubtitleBar: params.overlay.togglePrimarySubtitleBar,
      setVisible: params.overlay.setVisible,
    },
    mining: {
      copyCurrentSubtitle: params.mining.copyCurrentSubtitle,
      startPendingMultiCopy: params.mining.startPendingMultiCopy,
      mineSentenceCard: params.mining.mineSentenceCard,
      startPendingMineSentenceMultiple: params.mining.startPendingMineSentenceMultiple,
      updateLastCardFromClipboard: params.mining.updateLastCardFromClipboard,
      refreshKnownWords: params.mining.refreshKnownWords,
      triggerFieldGrouping: params.mining.triggerFieldGrouping,
      triggerSubsyncFromConfig: params.mining.triggerSubsyncFromConfig,
      markLastCardAsAudioCard: params.mining.markLastCardAsAudioCard,
    },
    anilist: {
      getStatus: params.anilist.getStatus,
      clearToken: params.anilist.clearToken,
      openSetup: params.anilist.openSetup,
      getQueueStatus: params.anilist.getQueueStatus,
      retryQueueNow: params.anilist.retryQueueNow,
    },
    dictionary: {
      generate: params.dictionary.generate,
      getSelection: params.dictionary.getSelection,
      setSelection: params.dictionary.setSelection,
    },
    jellyfin: {
      openSetup: params.jellyfin.openSetup,
      runStatsCommand: params.jellyfin.runStatsCommand,
      runCommand: params.jellyfin.runCommand,
    },
    app: {
      stop: params.app.stop,
      hasMainWindow: params.app.hasMainWindow,
      runUpdateCommand: params.app.runUpdateCommand,
      runEnsureLinuxRuntimePluginAssetsCommand: params.app.runEnsureLinuxRuntimePluginAssetsCommand,
      runYoutubePlaybackFlow: params.app.runYoutubePlaybackFlow,
    },
    dispatchSessionAction: params.dispatchSessionAction,
    ui: {
      openFirstRunSetup: params.ui.openFirstRunSetup,
      openYomitanSettings: params.ui.openYomitanSettings,
      openConfigSettingsWindow: params.ui.openConfigSettingsWindow,
      cycleSecondarySubMode: params.ui.cycleSecondarySubMode,
      openRuntimeOptionsPalette: params.ui.openRuntimeOptionsPalette,
      printHelp: params.ui.printHelp,
    },
    getMultiCopyTimeoutMs: params.getMultiCopyTimeoutMs,
    schedule: params.schedule,
    log: params.log,
    logDebug: params.logDebug,
    warn: params.warn,
    error: params.error,
  };
}

export function createMpvCommandRuntimeServiceDeps(
  params: MpvCommandRuntimeServiceDepsParams,
): HandleMpvCommandFromIpcOptions {
  return {
    specialCommands: params.specialCommands,
    triggerSubsyncFromConfig: params.triggerSubsyncFromConfig,
    openRuntimeOptionsPalette: params.openRuntimeOptionsPalette,
    openJimaku: params.openJimaku,
    openYoutubeTrackPicker: params.openYoutubeTrackPicker,
    openPlaylistBrowser: params.openPlaylistBrowser,
    runtimeOptionsCycle: params.runtimeOptionsCycle,
    showMpvOsd: params.showMpvOsd,
    showRawMpvOsd: params.showRawMpvOsd,
    showPlaybackFeedback: params.showPlaybackFeedback,
    mpvReplaySubtitle: params.mpvReplaySubtitle,
    mpvPlayNextSubtitle: params.mpvPlayNextSubtitle,
    mpvSendCommand: params.mpvSendCommand,
    resolveProxyCommandOsd: params.resolveProxyCommandOsd,
    isMpvConnected: params.isMpvConnected,
    hasRuntimeOptionsManager: params.hasRuntimeOptionsManager,
  };
}

import {
  RuntimeOptionId,
  RuntimeOptionValue,
  SubsyncManualPayload,
} from "../types";
import { SubsyncResolvedConfig } from "../subsync/utils";
import type { SubsyncRuntimeDeps } from "../core/services/subsync-runner-service";
import type { IpcDepsRuntimeOptions } from "../core/services/ipc-service";
import type { AnkiJimakuIpcRuntimeOptions } from "../core/services/anki-jimaku-service";
import type { CliCommandDepsRuntimeOptions } from "../core/services/cli-command-service";
import type { HandleMpvCommandFromIpcOptions } from "../core/services/ipc-command-service";
import {
  cycleRuntimeOptionFromIpcRuntimeService,
  setRuntimeOptionFromIpcRuntimeService,
} from "../core/services/runtime-options-ipc-service";
import { RuntimeOptionsManager } from "../runtime-options";

export interface RuntimeOptionsIpcDepsParams {
  getRuntimeOptionsManager: () => RuntimeOptionsManager | null;
  showMpvOsd: (text: string) => void;
}

export interface SubsyncRuntimeDepsParams {
  getMpvClient: () => ReturnType<SubsyncRuntimeDeps["getMpvClient"]>;
  getResolvedSubsyncConfig: () => SubsyncResolvedConfig;
  isSubsyncInProgress: () => boolean;
  setSubsyncInProgress: (inProgress: boolean) => void;
  showMpvOsd: (text: string) => void;
  openManualPicker: (payload: SubsyncManualPayload) => void;
}

export function createRuntimeOptionsIpcDeps(params: RuntimeOptionsIpcDepsParams): {
  setRuntimeOption: (id: string, value: unknown) => unknown;
  cycleRuntimeOption: (id: string, direction: 1 | -1) => unknown;
} {
  return {
    setRuntimeOption: (id, value) =>
      setRuntimeOptionFromIpcRuntimeService(
        params.getRuntimeOptionsManager(),
        id as RuntimeOptionId,
        value as RuntimeOptionValue,
        (text) => params.showMpvOsd(text),
      ),
    cycleRuntimeOption: (id, direction) =>
      cycleRuntimeOptionFromIpcRuntimeService(
        params.getRuntimeOptionsManager(),
        id as RuntimeOptionId,
        direction,
        (text) => params.showMpvOsd(text),
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
  getInvisibleWindow: IpcDepsRuntimeOptions["getInvisibleWindow"];
  getMainWindow: IpcDepsRuntimeOptions["getMainWindow"];
  getVisibleOverlayVisibility: IpcDepsRuntimeOptions["getVisibleOverlayVisibility"];
  getInvisibleOverlayVisibility: IpcDepsRuntimeOptions["getInvisibleOverlayVisibility"];
  onOverlayModalClosed: IpcDepsRuntimeOptions["onOverlayModalClosed"];
  openYomitanSettings: IpcDepsRuntimeOptions["openYomitanSettings"];
  quitApp: IpcDepsRuntimeOptions["quitApp"];
  toggleVisibleOverlay: IpcDepsRuntimeOptions["toggleVisibleOverlay"];
  tokenizeCurrentSubtitle: IpcDepsRuntimeOptions["tokenizeCurrentSubtitle"];
  getCurrentSubtitleAss: IpcDepsRuntimeOptions["getCurrentSubtitleAss"];
  focusMainWindow?: IpcDepsRuntimeOptions["focusMainWindow"];
  getMpvSubtitleRenderMetrics: IpcDepsRuntimeOptions["getMpvSubtitleRenderMetrics"];
  getSubtitlePosition: IpcDepsRuntimeOptions["getSubtitlePosition"];
  getSubtitleStyle: IpcDepsRuntimeOptions["getSubtitleStyle"];
  saveSubtitlePosition: IpcDepsRuntimeOptions["saveSubtitlePosition"];
  getMecabTokenizer: IpcDepsRuntimeOptions["getMecabTokenizer"];
  handleMpvCommand: IpcDepsRuntimeOptions["handleMpvCommand"];
  getKeybindings: IpcDepsRuntimeOptions["getKeybindings"];
  getConfiguredShortcuts: IpcDepsRuntimeOptions["getConfiguredShortcuts"];
  getSecondarySubMode: IpcDepsRuntimeOptions["getSecondarySubMode"];
  getMpvClient: IpcDepsRuntimeOptions["getMpvClient"];
  runSubsyncManual: IpcDepsRuntimeOptions["runSubsyncManual"];
  getAnkiConnectStatus: IpcDepsRuntimeOptions["getAnkiConnectStatus"];
  getRuntimeOptions: IpcDepsRuntimeOptions["getRuntimeOptions"];
  setRuntimeOption: IpcDepsRuntimeOptions["setRuntimeOption"];
  cycleRuntimeOption: IpcDepsRuntimeOptions["cycleRuntimeOption"];
  reportOverlayContentBounds: IpcDepsRuntimeOptions["reportOverlayContentBounds"];
}

export interface AnkiJimakuIpcRuntimeServiceDepsParams {
  patchAnkiConnectEnabled: AnkiJimakuIpcRuntimeOptions["patchAnkiConnectEnabled"];
  getResolvedConfig: AnkiJimakuIpcRuntimeOptions["getResolvedConfig"];
  getRuntimeOptionsManager: AnkiJimakuIpcRuntimeOptions["getRuntimeOptionsManager"];
  getSubtitleTimingTracker: AnkiJimakuIpcRuntimeOptions["getSubtitleTimingTracker"];
  getMpvClient: AnkiJimakuIpcRuntimeOptions["getMpvClient"];
  getAnkiIntegration: AnkiJimakuIpcRuntimeOptions["getAnkiIntegration"];
  setAnkiIntegration: AnkiJimakuIpcRuntimeOptions["setAnkiIntegration"];
  getKnownWordCacheStatePath: AnkiJimakuIpcRuntimeOptions["getKnownWordCacheStatePath"];
  showDesktopNotification: AnkiJimakuIpcRuntimeOptions["showDesktopNotification"];
  createFieldGroupingCallback: AnkiJimakuIpcRuntimeOptions["createFieldGroupingCallback"];
  broadcastRuntimeOptionsChanged: AnkiJimakuIpcRuntimeOptions["broadcastRuntimeOptionsChanged"];
  getFieldGroupingResolver: AnkiJimakuIpcRuntimeOptions["getFieldGroupingResolver"];
  setFieldGroupingResolver: AnkiJimakuIpcRuntimeOptions["setFieldGroupingResolver"];
  parseMediaInfo: AnkiJimakuIpcRuntimeOptions["parseMediaInfo"];
  getCurrentMediaPath: AnkiJimakuIpcRuntimeOptions["getCurrentMediaPath"];
  jimakuFetchJson: AnkiJimakuIpcRuntimeOptions["jimakuFetchJson"];
  getJimakuMaxEntryResults: AnkiJimakuIpcRuntimeOptions["getJimakuMaxEntryResults"];
  getJimakuLanguagePreference: AnkiJimakuIpcRuntimeOptions["getJimakuLanguagePreference"];
  resolveJimakuApiKey: AnkiJimakuIpcRuntimeOptions["resolveJimakuApiKey"];
  isRemoteMediaPath: AnkiJimakuIpcRuntimeOptions["isRemoteMediaPath"];
  downloadToFile: AnkiJimakuIpcRuntimeOptions["downloadToFile"];
}

export interface CliCommandRuntimeServiceDepsParams {
  mpv: {
    getSocketPath: CliCommandDepsRuntimeOptions["mpv"]["getSocketPath"];
    setSocketPath: CliCommandDepsRuntimeOptions["mpv"]["setSocketPath"];
    getClient: CliCommandDepsRuntimeOptions["mpv"]["getClient"];
    showOsd: CliCommandDepsRuntimeOptions["mpv"]["showOsd"];
  };
  texthooker: {
    service: CliCommandDepsRuntimeOptions["texthooker"]["service"];
    getPort: CliCommandDepsRuntimeOptions["texthooker"]["getPort"];
    setPort: CliCommandDepsRuntimeOptions["texthooker"]["setPort"];
    shouldOpenBrowser: CliCommandDepsRuntimeOptions["texthooker"]["shouldOpenBrowser"];
    openInBrowser: CliCommandDepsRuntimeOptions["texthooker"]["openInBrowser"];
  };
  overlay: {
    isInitialized: CliCommandDepsRuntimeOptions["overlay"]["isInitialized"];
    initialize: CliCommandDepsRuntimeOptions["overlay"]["initialize"];
    toggleVisible: CliCommandDepsRuntimeOptions["overlay"]["toggleVisible"];
    toggleInvisible: CliCommandDepsRuntimeOptions["overlay"]["toggleInvisible"];
    setVisible: CliCommandDepsRuntimeOptions["overlay"]["setVisible"];
    setInvisible: CliCommandDepsRuntimeOptions["overlay"]["setInvisible"];
  };
  mining: {
    copyCurrentSubtitle: CliCommandDepsRuntimeOptions["mining"]["copyCurrentSubtitle"];
    startPendingMultiCopy:
      CliCommandDepsRuntimeOptions["mining"]["startPendingMultiCopy"];
    mineSentenceCard: CliCommandDepsRuntimeOptions["mining"]["mineSentenceCard"];
    startPendingMineSentenceMultiple:
      CliCommandDepsRuntimeOptions["mining"]["startPendingMineSentenceMultiple"];
    updateLastCardFromClipboard:
      CliCommandDepsRuntimeOptions["mining"]["updateLastCardFromClipboard"];
    refreshKnownWords: CliCommandDepsRuntimeOptions["mining"]["refreshKnownWords"];
    triggerFieldGrouping: CliCommandDepsRuntimeOptions["mining"]["triggerFieldGrouping"];
    triggerSubsyncFromConfig:
      CliCommandDepsRuntimeOptions["mining"]["triggerSubsyncFromConfig"];
    markLastCardAsAudioCard:
      CliCommandDepsRuntimeOptions["mining"]["markLastCardAsAudioCard"];
  };
  ui: {
    openYomitanSettings: CliCommandDepsRuntimeOptions["ui"]["openYomitanSettings"];
    cycleSecondarySubMode: CliCommandDepsRuntimeOptions["ui"]["cycleSecondarySubMode"];
    openRuntimeOptionsPalette:
      CliCommandDepsRuntimeOptions["ui"]["openRuntimeOptionsPalette"];
    printHelp: CliCommandDepsRuntimeOptions["ui"]["printHelp"];
  };
  app: {
    stop: CliCommandDepsRuntimeOptions["app"]["stop"];
    hasMainWindow: CliCommandDepsRuntimeOptions["app"]["hasMainWindow"];
  };
  getMultiCopyTimeoutMs: CliCommandDepsRuntimeOptions["getMultiCopyTimeoutMs"];
  schedule: CliCommandDepsRuntimeOptions["schedule"];
  log: CliCommandDepsRuntimeOptions["log"];
  warn: CliCommandDepsRuntimeOptions["warn"];
  error: CliCommandDepsRuntimeOptions["error"];
}

export interface MpvCommandRuntimeServiceDepsParams {
  specialCommands: HandleMpvCommandFromIpcOptions["specialCommands"];
  runtimeOptionsCycle: HandleMpvCommandFromIpcOptions["runtimeOptionsCycle"];
  triggerSubsyncFromConfig: HandleMpvCommandFromIpcOptions["triggerSubsyncFromConfig"];
  openRuntimeOptionsPalette: HandleMpvCommandFromIpcOptions["openRuntimeOptionsPalette"];
  showMpvOsd: HandleMpvCommandFromIpcOptions["showMpvOsd"];
  mpvReplaySubtitle: HandleMpvCommandFromIpcOptions["mpvReplaySubtitle"];
  mpvPlayNextSubtitle: HandleMpvCommandFromIpcOptions["mpvPlayNextSubtitle"];
  mpvSendCommand: HandleMpvCommandFromIpcOptions["mpvSendCommand"];
  isMpvConnected: HandleMpvCommandFromIpcOptions["isMpvConnected"];
  hasRuntimeOptionsManager: HandleMpvCommandFromIpcOptions["hasRuntimeOptionsManager"];
}

export function createMainIpcRuntimeServiceDeps(
  params: MainIpcRuntimeServiceDepsParams,
): IpcDepsRuntimeOptions {
  return {
    getInvisibleWindow: params.getInvisibleWindow,
    getMainWindow: params.getMainWindow,
    getVisibleOverlayVisibility: params.getVisibleOverlayVisibility,
    getInvisibleOverlayVisibility: params.getInvisibleOverlayVisibility,
    onOverlayModalClosed: params.onOverlayModalClosed,
    openYomitanSettings: params.openYomitanSettings,
    quitApp: params.quitApp,
    toggleVisibleOverlay: params.toggleVisibleOverlay,
    tokenizeCurrentSubtitle: params.tokenizeCurrentSubtitle,
    getCurrentSubtitleAss: params.getCurrentSubtitleAss,
    getMpvSubtitleRenderMetrics: params.getMpvSubtitleRenderMetrics,
    getSubtitlePosition: params.getSubtitlePosition,
    getSubtitleStyle: params.getSubtitleStyle,
    saveSubtitlePosition: params.saveSubtitlePosition,
    getMecabTokenizer: params.getMecabTokenizer,
    handleMpvCommand: params.handleMpvCommand,
    getKeybindings: params.getKeybindings,
    getConfiguredShortcuts: params.getConfiguredShortcuts,
    focusMainWindow: params.focusMainWindow ?? (() => {}),
    getSecondarySubMode: params.getSecondarySubMode,
    getMpvClient: params.getMpvClient,
    runSubsyncManual: params.runSubsyncManual,
    getAnkiConnectStatus: params.getAnkiConnectStatus,
    getRuntimeOptions: params.getRuntimeOptions,
    setRuntimeOption: params.setRuntimeOption,
    cycleRuntimeOption: params.cycleRuntimeOption,
    reportOverlayContentBounds: params.reportOverlayContentBounds,
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
    showDesktopNotification: params.showDesktopNotification,
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
    mpv: {
      getSocketPath: params.mpv.getSocketPath,
      setSocketPath: params.mpv.setSocketPath,
      getClient: params.mpv.getClient,
      showOsd: params.mpv.showOsd,
    },
    texthooker: {
      service: params.texthooker.service,
      getPort: params.texthooker.getPort,
      setPort: params.texthooker.setPort,
      shouldOpenBrowser: params.texthooker.shouldOpenBrowser,
      openInBrowser: params.texthooker.openInBrowser,
    },
    overlay: {
      isInitialized: params.overlay.isInitialized,
      initialize: params.overlay.initialize,
      toggleVisible: params.overlay.toggleVisible,
      toggleInvisible: params.overlay.toggleInvisible,
      setVisible: params.overlay.setVisible,
      setInvisible: params.overlay.setInvisible,
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
    ui: {
      openYomitanSettings: params.ui.openYomitanSettings,
      cycleSecondarySubMode: params.ui.cycleSecondarySubMode,
      openRuntimeOptionsPalette: params.ui.openRuntimeOptionsPalette,
      printHelp: params.ui.printHelp,
    },
    app: {
      stop: params.app.stop,
      hasMainWindow: params.app.hasMainWindow,
    },
    getMultiCopyTimeoutMs: params.getMultiCopyTimeoutMs,
    schedule: params.schedule,
    log: params.log,
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
    runtimeOptionsCycle: params.runtimeOptionsCycle,
    showMpvOsd: params.showMpvOsd,
    mpvReplaySubtitle: params.mpvReplaySubtitle,
    mpvPlayNextSubtitle: params.mpvPlayNextSubtitle,
    mpvSendCommand: params.mpvSendCommand,
    isMpvConnected: params.isMpvConnected,
    hasRuntimeOptionsManager: params.hasRuntimeOptionsManager,
  };
}

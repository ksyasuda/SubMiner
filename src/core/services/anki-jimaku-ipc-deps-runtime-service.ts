import {
  AnkiJimakuIpcRuntimeOptions,
} from "./anki-jimaku-runtime-service";

export type AnkiJimakuIpcDepsRuntimeOptions = AnkiJimakuIpcRuntimeOptions;

export function createAnkiJimakuIpcDepsRuntimeService(
  options: AnkiJimakuIpcDepsRuntimeOptions,
): AnkiJimakuIpcRuntimeOptions {
  return {
    patchAnkiConnectEnabled: options.patchAnkiConnectEnabled,
    getResolvedConfig: options.getResolvedConfig,
    getRuntimeOptionsManager: options.getRuntimeOptionsManager,
    getSubtitleTimingTracker: options.getSubtitleTimingTracker,
    getMpvClient: options.getMpvClient,
    getAnkiIntegration: options.getAnkiIntegration,
    setAnkiIntegration: options.setAnkiIntegration,
    showDesktopNotification: options.showDesktopNotification,
    createFieldGroupingCallback: options.createFieldGroupingCallback,
    broadcastRuntimeOptionsChanged: options.broadcastRuntimeOptionsChanged,
    getFieldGroupingResolver: options.getFieldGroupingResolver,
    setFieldGroupingResolver: options.setFieldGroupingResolver,
    parseMediaInfo: options.parseMediaInfo,
    getCurrentMediaPath: options.getCurrentMediaPath,
    jimakuFetchJson: options.jimakuFetchJson,
    getJimakuMaxEntryResults: options.getJimakuMaxEntryResults,
    getJimakuLanguagePreference: options.getJimakuLanguagePreference,
    resolveJimakuApiKey: options.resolveJimakuApiKey,
    isRemoteMediaPath: options.isRemoteMediaPath,
    downloadToFile: options.downloadToFile,
  };
}

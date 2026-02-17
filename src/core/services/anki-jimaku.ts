import { AnkiIntegration } from "../../anki-integration";
import {
  AnkiConnectConfig,
  JimakuApiResponse,
  JimakuEntry,
  JimakuFileEntry,
  JimakuLanguagePreference,
  JimakuMediaInfo,
  KikuFieldGroupingChoice,
  KikuFieldGroupingRequestData,
} from "../../types";
import { sortJimakuFiles } from "../../jimaku/utils";
import type { AnkiJimakuIpcDeps } from "./anki-jimaku-ipc";
import { createLogger } from "../../logger";

export type RegisterAnkiJimakuIpcRuntimeHandler = (
  deps: AnkiJimakuIpcDeps,
) => void;

interface MpvClientLike {
  connected: boolean;
  send: (payload: { command: string[] }) => void;
}

interface RuntimeOptionsManagerLike {
  getEffectiveAnkiConnectConfig: (config?: AnkiConnectConfig) => AnkiConnectConfig;
}

interface SubtitleTimingTrackerLike {
  cleanup: () => void;
}

export interface AnkiJimakuIpcRuntimeOptions {
  patchAnkiConnectEnabled: (enabled: boolean) => void;
  getResolvedConfig: () => { ankiConnect?: AnkiConnectConfig };
  getRuntimeOptionsManager: () => RuntimeOptionsManagerLike | null;
  getSubtitleTimingTracker: () => SubtitleTimingTrackerLike | null;
  getMpvClient: () => MpvClientLike | null;
  getAnkiIntegration: () => AnkiIntegration | null;
  setAnkiIntegration: (integration: AnkiIntegration | null) => void;
  getKnownWordCacheStatePath: () => string;
  showDesktopNotification: (title: string, options: { body?: string; icon?: string }) => void;
  createFieldGroupingCallback: () => (
    data: KikuFieldGroupingRequestData,
  ) => Promise<KikuFieldGroupingChoice>;
  broadcastRuntimeOptionsChanged: () => void;
  getFieldGroupingResolver: () => ((choice: KikuFieldGroupingChoice) => void) | null;
  setFieldGroupingResolver: (resolver: ((choice: KikuFieldGroupingChoice) => void) | null) => void;
  parseMediaInfo: (mediaPath: string | null) => JimakuMediaInfo;
  getCurrentMediaPath: () => string | null;
  jimakuFetchJson: <T>(
    endpoint: string,
    query?: Record<string, string | number | boolean | null | undefined>,
  ) => Promise<JimakuApiResponse<T>>;
  getJimakuMaxEntryResults: () => number;
  getJimakuLanguagePreference: () => JimakuLanguagePreference;
  resolveJimakuApiKey: () => Promise<string | null>;
  isRemoteMediaPath: (mediaPath: string) => boolean;
  downloadToFile: (
    url: string,
    destPath: string,
    headers: Record<string, string>,
  ) => Promise<{ ok: true; path: string } | { ok: false; error: { error: string; code?: number; retryAfter?: number } }>;
}

const logger = createLogger("main:anki-jimaku");

export function registerAnkiJimakuIpcRuntime(
  options: AnkiJimakuIpcRuntimeOptions,
  registerHandlers: RegisterAnkiJimakuIpcRuntimeHandler,
): void {
  registerHandlers({
    setAnkiConnectEnabled: (enabled) => {
      options.patchAnkiConnectEnabled(enabled);
      const config = options.getResolvedConfig();
      const subtitleTimingTracker = options.getSubtitleTimingTracker();
      const mpvClient = options.getMpvClient();
      const ankiIntegration = options.getAnkiIntegration();

      if (enabled && !ankiIntegration && subtitleTimingTracker && mpvClient) {
        const runtimeOptionsManager = options.getRuntimeOptionsManager();
        const effectiveAnkiConfig = runtimeOptionsManager
          ? runtimeOptionsManager.getEffectiveAnkiConnectConfig(config.ankiConnect)
          : config.ankiConnect;
        const integration = new AnkiIntegration(
          effectiveAnkiConfig as never,
          subtitleTimingTracker as never,
          mpvClient as never,
          (text: string) => {
            if (mpvClient) {
              mpvClient.send({
                command: ["show-text", text, "3000"],
              });
            }
          },
          options.showDesktopNotification,
          options.createFieldGroupingCallback(),
          options.getKnownWordCacheStatePath(),
        );
        integration.start();
        options.setAnkiIntegration(integration);
        logger.info("AnkiConnect integration enabled");
      } else if (!enabled && ankiIntegration) {
        ankiIntegration.destroy();
        options.setAnkiIntegration(null);
        logger.info("AnkiConnect integration disabled");
      }

      options.broadcastRuntimeOptionsChanged();
    },
    clearAnkiHistory: () => {
      const subtitleTimingTracker = options.getSubtitleTimingTracker();
      if (subtitleTimingTracker) {
        subtitleTimingTracker.cleanup();
        logger.info("AnkiConnect subtitle timing history cleared");
      }
    },
    refreshKnownWords: async () => {
      const integration = options.getAnkiIntegration();
      if (!integration) {
        throw new Error("AnkiConnect integration not enabled");
      }
      await integration.refreshKnownWordCache();
    },
    respondFieldGrouping: (choice) => {
      const resolver = options.getFieldGroupingResolver();
      if (resolver) {
        resolver(choice);
        options.setFieldGroupingResolver(null);
      }
    },
    buildKikuMergePreview: async (request) => {
      const integration = options.getAnkiIntegration();
      if (!integration) {
        return { ok: false, error: "AnkiConnect integration not enabled" };
      }
      return integration.buildFieldGroupingPreview(
        request.keepNoteId,
        request.deleteNoteId,
        request.deleteDuplicate,
      );
    },
    getJimakuMediaInfo: () => options.parseMediaInfo(options.getCurrentMediaPath()),
    searchJimakuEntries: async (query) => {
      logger.info(`[jimaku] search-entries query: "${query.query}"`);
      const response = await options.jimakuFetchJson<JimakuEntry[]>(
        "/api/entries/search",
        {
          anime: true,
          query: query.query,
        },
      );
      if (!response.ok) return response;
      const maxResults = options.getJimakuMaxEntryResults();
      logger.info(
        `[jimaku] search-entries returned ${response.data.length} results (capped to ${maxResults})`,
      );
      return { ok: true, data: response.data.slice(0, maxResults) };
    },
    listJimakuFiles: async (query) => {
      logger.info(
        `[jimaku] list-files entryId=${query.entryId} episode=${query.episode ?? "all"}`,
      );
      const response = await options.jimakuFetchJson<JimakuFileEntry[]>(
        `/api/entries/${query.entryId}/files`,
        {
          episode: query.episode ?? undefined,
        },
      );
      if (!response.ok) return response;
      const sorted = sortJimakuFiles(
        response.data,
        options.getJimakuLanguagePreference(),
      );
      logger.info(`[jimaku] list-files returned ${sorted.length} files`);
      return { ok: true, data: sorted };
    },
    resolveJimakuApiKey: () => options.resolveJimakuApiKey(),
    getCurrentMediaPath: () => options.getCurrentMediaPath(),
    isRemoteMediaPath: (mediaPath) => options.isRemoteMediaPath(mediaPath),
    downloadToFile: (url, destPath, headers) =>
      options.downloadToFile(url, destPath, headers),
    onDownloadedSubtitle: (pathToSubtitle) => {
      const mpvClient = options.getMpvClient();
      if (mpvClient && mpvClient.connected) {
        mpvClient.send({ command: ["sub-add", pathToSubtitle, "select"] });
      }
    },
  });
}

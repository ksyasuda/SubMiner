import * as fs from 'fs';
import { AnkiIntegration } from '../../anki-integration';
import { mergeAiConfig } from '../../ai/config';
import {
  AiConfig,
  AnimetoshoApiResponse,
  AnimetoshoConfig,
  AnkiConnectConfig,
  JimakuApiResponse,
  JimakuEntry,
  JimakuFileEntry,
  JimakuLanguagePreference,
  JimakuMediaInfo,
  KikuFieldGroupingChoice,
  KikuFieldGroupingRequestData,
  OverlayNotificationPayload,
} from '../../types';
import { sortJimakuFiles } from '../../jimaku/utils';
import {
  TSUKIHIME_API_BASE_URL,
  animetoshoFetchJson as animetoshoFetchJsonRequest,
  decompressXzFile,
  extractAnimetoshoSubtitleFiles,
  isAnimetoshoDownloadUrl,
  mapAnimetoshoSearchResults,
} from '../../animetosho/utils';
import type { AnkiJimakuIpcDeps } from './anki-jimaku-ipc';
import { createLogger } from '../../logger';

export type RegisterAnkiJimakuIpcRuntimeHandler = (deps: AnkiJimakuIpcDeps) => void;

interface MpvClientLike {
  connected: boolean;
  send: (payload: { command: (string | number)[] }) => void;
  request?: (command: unknown[]) => Promise<{ data?: unknown }>;
}

interface RuntimeOptionsManagerLike {
  getEffectiveAnkiConnectConfig: (config?: AnkiConnectConfig) => AnkiConnectConfig;
}

interface SubtitleTimingTrackerLike {
  cleanup: () => void;
}

export interface AnkiJimakuIpcRuntimeOptions {
  patchAnkiConnectEnabled: (enabled: boolean) => void;
  getResolvedConfig: () => {
    ankiConnect?: AnkiConnectConfig;
    ai?: AiConfig;
    animetosho?: AnimetoshoConfig;
    secondarySub?: { secondarySubLanguages?: string[] };
  };
  getRuntimeOptionsManager: () => RuntimeOptionsManagerLike | null;
  getSubtitleTimingTracker: () => SubtitleTimingTrackerLike | null;
  getMpvClient: () => MpvClientLike | null;
  getAnkiIntegration: () => AnkiIntegration | null;
  setAnkiIntegration: (integration: AnkiIntegration | null) => void;
  getKnownWordCacheStatePath: () => string;
  getCachedMediaPath?: (
    currentVideoPath: string,
    kind: 'audio' | 'video',
  ) => Promise<string | null>;
  shouldRequireRemoteMediaCache?: () => boolean;
  getYoutubeMediaSourceUrl?: () => Promise<string | null | undefined> | string | null | undefined;
  showDesktopNotification: (title: string, options: { body?: string; icon?: string }) => void;
  showOverlayNotification?: (payload: OverlayNotificationPayload) => void;
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
  animetoshoFetchJson?: <T>(
    endpoint: string,
    query?: Record<string, string | number | boolean | null | undefined>,
  ) => Promise<AnimetoshoApiResponse<T>>;
  getJimakuMaxEntryResults: () => number;
  getJimakuLanguagePreference: () => JimakuLanguagePreference;
  resolveJimakuApiKey: () => Promise<string | null>;
  isRemoteMediaPath: (mediaPath: string) => boolean;
  downloadToFile: (
    url: string,
    destPath: string,
    headers: Record<string, string>,
    downloadOptions?: { isAllowedRedirect?: (url: URL) => boolean },
  ) => Promise<
    | { ok: true; path: string }
    | {
        ok: false;
        error: { error: string; code?: number; retryAfter?: number };
      }
  >;
}

const logger = createLogger('main:anki-jimaku');

const DEFAULT_ANIMETOSHO_MAX_SEARCH_RESULTS = 10;
const SECONDARY_TRACK_LOOKUP_ATTEMPTS = 5;
const SECONDARY_TRACK_LOOKUP_RETRY_MS = 100;

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function getAnimetoshoMaxSearchResults(options: AnkiJimakuIpcRuntimeOptions): number {
  const value = options.getResolvedConfig().animetosho?.maxSearchResults;
  if (typeof value === 'number' && Number.isFinite(value) && value > 0) {
    return Math.floor(value);
  }
  return DEFAULT_ANIMETOSHO_MAX_SEARCH_RESULTS;
}

function animetoshoFetch<T>(
  options: AnkiJimakuIpcRuntimeOptions,
  endpoint: string,
  query: Record<string, string | number | boolean | null | undefined>,
): Promise<AnimetoshoApiResponse<T>> {
  if (options.animetoshoFetchJson) {
    return options.animetoshoFetchJson<T>(endpoint, query);
  }
  const baseUrl = options.getResolvedConfig().animetosho?.apiBaseUrl || TSUKIHIME_API_BASE_URL;
  return animetoshoFetchJsonRequest<T>(endpoint, query, { baseUrl });
}

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
                command: ['show-text', text, '3000'],
              });
            }
          },
          options.showDesktopNotification,
          options.createFieldGroupingCallback(),
          options.getKnownWordCacheStatePath(),
          mergeAiConfig(config.ai, config.ankiConnect?.ai) as AiConfig,
          undefined,
          options.showOverlayNotification,
          options.getCachedMediaPath,
          options.shouldRequireRemoteMediaCache,
          options.getYoutubeMediaSourceUrl,
        );
        integration.start();
        options.setAnkiIntegration(integration);
        logger.info('AnkiConnect integration enabled');
      } else if (!enabled && ankiIntegration) {
        ankiIntegration.destroy();
        options.setAnkiIntegration(null);
        logger.info('AnkiConnect integration disabled');
      }

      options.broadcastRuntimeOptionsChanged();
    },
    clearAnkiHistory: () => {
      const subtitleTimingTracker = options.getSubtitleTimingTracker();
      if (subtitleTimingTracker) {
        subtitleTimingTracker.cleanup();
        logger.info('AnkiConnect subtitle timing history cleared');
      }
    },
    refreshKnownWords: async () => {
      const integration = options.getAnkiIntegration();
      if (!integration) {
        throw new Error('AnkiConnect integration not enabled');
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
        return { ok: false, error: 'AnkiConnect integration not enabled' };
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
      const response = await options.jimakuFetchJson<JimakuEntry[]>('/api/entries/search', {
        anime: true,
        query: query.query,
      });
      if (!response.ok) return response;
      const maxResults = options.getJimakuMaxEntryResults();
      logger.info(
        `[jimaku] search-entries returned ${response.data.length} results (capped to ${maxResults})`,
      );
      return { ok: true, data: response.data.slice(0, maxResults) };
    },
    listJimakuFiles: async (query) => {
      logger.info(`[jimaku] list-files entryId=${query.entryId} episode=${query.episode ?? 'all'}`);
      const response = await options.jimakuFetchJson<JimakuFileEntry[]>(
        `/api/entries/${query.entryId}/files`,
        {
          episode: query.episode ?? undefined,
        },
      );
      if (!response.ok) return response;
      const sorted = sortJimakuFiles(response.data, options.getJimakuLanguagePreference());
      logger.info(`[jimaku] list-files returned ${sorted.length} files`);
      return { ok: true, data: sorted };
    },
    resolveJimakuApiKey: () => options.resolveJimakuApiKey(),
    getCurrentMediaPath: () => options.getCurrentMediaPath(),
    isRemoteMediaPath: (mediaPath) => options.isRemoteMediaPath(mediaPath),
    downloadToFile: (url, destPath, headers) => options.downloadToFile(url, destPath, headers),

    searchAnimetoshoEntries: async (query) => {
      logger.info(`[animetosho] search-entries query: "${query.query}"`);
      const maxResults = getAnimetoshoMaxSearchResults(options);
      const response = await animetoshoFetch<unknown>(options, '/search/torrents', {
        q: query.query,
        // The API caps limit at 100.
        limit: Math.min(maxResults, 100),
      });
      if (!response.ok) return response;
      const entries = mapAnimetoshoSearchResults(response.data, maxResults);
      logger.info(`[animetosho] search-entries returned ${entries.length} results`);
      return { ok: true, data: entries };
    },
    listAnimetoshoFiles: async (query) => {
      logger.info(`[animetosho] list-files entryId=${query.entryId}`);
      const response = await animetoshoFetch<unknown>(
        options,
        `/torrents/${encodeURIComponent(query.entryId)}`,
        {},
      );
      if (!response.ok) return response;
      const files = extractAnimetoshoSubtitleFiles(response.data);
      logger.info(`[animetosho] list-files returned ${files.length} subtitle attachments`);
      return { ok: true, data: files };
    },
    getAnimetoshoSecondaryLanguages: () =>
      options.getResolvedConfig().secondarySub?.secondarySubLanguages ?? [],
    downloadAnimetoshoSubtitle: async (url, destPath) => {
      const tempXzPath = `${destPath}.xz`;
      const downloaded = await options.downloadToFile(
        url,
        tempXzPath,
        { 'User-Agent': 'SubMiner' },
        // animetosho.org redirects to storage.animetosho.org; keep the hop in-domain.
        { isAllowedRedirect: (redirectUrl) => isAnimetoshoDownloadUrl(redirectUrl) },
      );
      if (!downloaded.ok) return downloaded;
      const result = await decompressXzFile(tempXzPath, destPath);
      fs.promises.unlink(tempXzPath).catch(() => {});
      return result;
    },
    onDownloadedSubtitle: (pathToSubtitle) => {
      const mpvClient = options.getMpvClient();
      if (mpvClient && mpvClient.connected) {
        mpvClient.send({ command: ['sub-add', pathToSubtitle, 'select'] });
      }
    },
    onDownloadedSecondarySubtitle: async (pathToSubtitle) => {
      const mpvClient = options.getMpvClient();
      if (!mpvClient || !mpvClient.connected) return;
      mpvClient.send({ command: ['sub-add', pathToSubtitle, 'auto'] });
      const request = mpvClient.request;
      if (!request) return;

      // sub-add is queued, so the track may not appear in the first track-list
      // reply; poll briefly before giving up.
      for (let attempt = 0; attempt < SECONDARY_TRACK_LOOKUP_ATTEMPTS; attempt += 1) {
        try {
          const response = await request(['get_property', 'track-list']);
          const tracks = Array.isArray(response?.data)
            ? (response.data as Array<Record<string, unknown>>)
            : [];
          const added = tracks.find(
            (track) => track?.type === 'sub' && track['external-filename'] === pathToSubtitle,
          );
          if (added && typeof added.id === 'number') {
            mpvClient.send({ command: ['set_property', 'secondary-sid', added.id] });
            return;
          }
        } catch (error) {
          logger.warn('[animetosho] failed to select downloaded subtitle as secondary:', error);
          return;
        }
        await delay(SECONDARY_TRACK_LOOKUP_RETRY_MS);
      }

      logger.warn(
        `[animetosho] could not find downloaded subtitle in track-list: ${pathToSubtitle}`,
      );
    },
  });
}

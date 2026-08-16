import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import type { AnimeBrowserRuntime } from './anime-browser-runtime';
import type {
  AnimeBrowserPlayRequest,
  AnimeBrowserSetWatchedRequest,
  AnimeBrowserWatchStateRequest,
} from '../../types/anime-browser';

export interface AnimeBrowserIpcDeps {
  // Structurally typed so tests can pass a fake without importing Electron.
  ipcMain: {
    handle(channel: string, listener: (event: unknown, ...args: unknown[]) => unknown): unknown;
  };
  runtime: AnimeBrowserRuntime;
  getPlaybackState?: () => unknown;
  registerSession?: (sessionId: string, sender: AnimeBrowserIpcSender) => void;
}

export interface AnimeBrowserIpcSender {
  send(channel: string, ...args: unknown[]): void;
  isDestroyed(): boolean;
  once(event: 'destroyed', listener: () => void): unknown;
}

interface AnimeBrowserIpcEvent {
  sender: AnimeBrowserIpcSender;
}

/**
 * Bridge the renderer to the anime runtime. Page arguments arrive as `unknown`
 * from the renderer, so they are coerced here rather than trusted.
 */
export function registerAnimeBrowserIpcHandlers(deps: AnimeBrowserIpcDeps): void {
  const channels = IPC_CHANNELS.request;
  const { runtime } = deps;
  const handle = (
    channel: string,
    listener: (event: unknown, ...args: unknown[]) => unknown,
  ): void => {
    deps.ipcMain.handle(channel, listener);
  };

  handle(channels.animeBrowserGetSnapshot, (event, sessionId) => {
    const session = registerSession(deps, event, sessionId);
    return runtime.getSnapshot(session);
  });
  handle(channels.animeBrowserEnsureBridge, (event, sessionId) => {
    registerSession(deps, event, sessionId);
    return runtime.ensureBridge();
  });
  handle(channels.animeBrowserSelectSource, (event, sessionId, sourceId) =>
    runtime.selectSource(String(sourceId), registerSession(deps, event, sessionId)),
  );
  handle(channels.animeBrowserSearch, (event, sessionId, query, page) =>
    runtime.search(String(query ?? ''), toPage(page), registerSession(deps, event, sessionId)),
  );
  handle(channels.animeBrowserGetPopular, (event, sessionId, page) =>
    runtime.getPopular(toPage(page), registerSession(deps, event, sessionId)),
  );
  handle(channels.animeBrowserGetDetails, (event, sessionId, animeUrl, sourceId) =>
    runtime.getDetails(
      String(animeUrl),
      toOptionalId(sourceId),
      registerSession(deps, event, sessionId),
    ),
  );
  handle(channels.animeBrowserGetEpisodes, (event, sessionId, animeUrl, sourceId) =>
    runtime.getEpisodes(
      String(animeUrl),
      toOptionalId(sourceId),
      registerSession(deps, event, sessionId),
    ),
  );
  handle(channels.animeBrowserGetWatchState, (_event, request) =>
    runtime.getWatchState(toWatchStateRequest(request)),
  );
  handle(channels.animeBrowserSetWatched, (_event, request) =>
    runtime.setWatched(toSetWatchedRequest(request)),
  );
  handle(channels.animeBrowserListAvailableExtensions, () => runtime.listAvailableExtensions());
  handle(channels.animeBrowserInstallExtension, (_event, pkg) =>
    runtime.installExtension(String(pkg)),
  );
  handle(channels.animeBrowserRemoveExtension, (_event, pkg) =>
    runtime.removeExtension(String(pkg)),
  );
  handle(channels.animeBrowserRescanExtensions, () => runtime.rescanExtensions());
  handle(channels.animeBrowserAddRepo, (_event, url) => runtime.addRepo(String(url)));
  handle(channels.animeBrowserRemoveRepo, (_event, url) => runtime.removeRepo(String(url)));
  handle(channels.animeBrowserPlayEpisode, (_event, request) =>
    runtime.playEpisode(toPlayRequest(request)),
  );
  handle(channels.animeBrowserQueueEpisode, (_event, request) =>
    runtime.queueEpisode(toPlayRequest(request)),
  );
  handle(channels.animeBrowserDequeueEpisode, (_event, sourceId, episodeUrl) =>
    runtime.dequeueEpisode(String(sourceId ?? ''), String(episodeUrl ?? '')),
  );
  handle(channels.animeBrowserClearQueue, () => runtime.clearQueue());
  handle(channels.animeBrowserGetQueue, () => runtime.getQueue());
  handle(channels.animeBrowserGetPlaybackState, () => deps.getPlaybackState?.() ?? null);
  handle(channels.animeBrowserIsPlaying, () => runtime.isPlaying());
  handle(channels.animeBrowserGetPreferences, (_event, sourceId) =>
    runtime.getPreferences(String(sourceId)),
  );
  handle(channels.animeBrowserSetPreference, (_event, sourceId, key, value) =>
    runtime.setPreference(String(sourceId), String(key), value as string | string[] | boolean),
  );
}

function registerSession(deps: AnimeBrowserIpcDeps, event: unknown, sessionId: unknown): string {
  const normalized = typeof sessionId === 'string' && sessionId.length > 0 ? sessionId : 'default';
  const sender = (event as AnimeBrowserIpcEvent | undefined)?.sender;
  if (sender) deps.registerSession?.(normalized, sender);
  return normalized;
}

/** Coerce a play (or queue) request at the renderer trust boundary. */
function toPlayRequest(value: unknown): AnimeBrowserPlayRequest {
  const request = (value ?? {}) as Partial<AnimeBrowserPlayRequest>;
  return {
    sourceId: String(request.sourceId ?? ''),
    animeUrl: String(request.animeUrl ?? ''),
    animeTitle: String(request.animeTitle ?? ''),
    episodeUrl: String(request.episodeUrl ?? ''),
    episodeName: String(request.episodeName ?? ''),
    // NaN and Infinity are numbers as far as typeof is concerned, and either
    // one would reach the stats row as a nonsense episode number.
    episodeNumber: Number.isFinite(request.episodeNumber) ? request.episodeNumber! : null,
  };
}

/** Coerce a watch-state request; the renderer's arrays arrive untyped. */
function toWatchStateRequest(value: unknown): AnimeBrowserWatchStateRequest {
  const request = (value ?? {}) as Partial<AnimeBrowserWatchStateRequest>;
  return {
    sourceId: String(request.sourceId ?? ''),
    animeUrl: String(request.animeUrl ?? ''),
    episodeUrls: Array.isArray(request.episodeUrls)
      ? request.episodeUrls.filter((url): url is string => typeof url === 'string')
      : [],
  };
}

/** Coerce a manual watch-mark request, including its per-episode entries. */
function toSetWatchedRequest(value: unknown): AnimeBrowserSetWatchedRequest {
  const request = (value ?? {}) as Partial<AnimeBrowserSetWatchedRequest>;
  const episodes = Array.isArray(request.episodes) ? request.episodes : [];
  return {
    sourceId: String(request.sourceId ?? ''),
    animeUrl: String(request.animeUrl ?? ''),
    animeTitle: String(request.animeTitle ?? ''),
    episodes: episodes.map((episode) => ({
      episodeUrl: String(episode?.episodeUrl ?? ''),
      episodeName: String(episode?.episodeName ?? ''),
      // NaN and Infinity are numbers as far as typeof is concerned, and either
      // one would reach the stats row as a nonsense episode number.
      episodeNumber: Number.isFinite(episode?.episodeNumber) ? episode.episodeNumber! : null,
    })),
    watched: request.watched === true,
  };
}

/**
 * A source id the renderer may omit. Absent means "use the current selection",
 * so an empty value must stay undefined rather than becoming the string "".
 */
function toOptionalId(value: unknown): string | undefined {
  return typeof value === 'string' && value.length > 0 ? value : undefined;
}

/** Bridge pages are 1-based; anything unusable falls back to the first page. */
function toPage(value: unknown): number {
  const page = Number(value);
  return Number.isFinite(page) && page >= 1 ? Math.floor(page) : 1;
}

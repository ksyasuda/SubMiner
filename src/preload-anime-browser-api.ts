import type { IpcRenderer } from 'electron';
import { IPC_CHANNELS } from './shared/ipc/contracts';
import type {
  AnimeBrowserAPI,
  AnimeBrowserBridgeState,
  AnimeBrowserDetails,
  AnimeBrowserEpisode,
  AnimeBrowserEpisodeWatchState,
  AnimeBrowserPlayRequest,
  AnimeBrowserPlayResult,
  AnimeBrowserPlaybackState,
  AnimeBrowserQueueState,
  AnimeBrowserSearchResult,
  AnimeBrowserSearchUpdate,
  AnimeBrowserSetWatchedRequest,
  AnimeBrowserSnapshot,
  AnimeBrowserWatchStateRequest,
  AvailableExtensionsResult,
  SourcePreferenceView,
} from './types/anime-browser';

type AnimeBrowserIpcRenderer = Pick<IpcRenderer, 'invoke' | 'on' | 'removeListener'>;

/**
 * Build the browser bridge for one renderer. The opaque session id keeps
 * source selection and in-flight searches independent between the standalone
 * window and the player-hosted modal.
 */
export function createAnimeBrowserAPI(ipcRenderer: AnimeBrowserIpcRenderer): AnimeBrowserAPI {
  const request = IPC_CHANNELS.request;
  // Sandboxed Electron preloads cannot require arbitrary Node built-ins.
  // Web Crypto is available in the isolated renderer world; the fallback only
  // covers test or legacy contexts where randomUUID is absent.
  const sessionId =
    globalThis.crypto?.randomUUID?.() ??
    `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;

  return {
    getSnapshot: (): Promise<AnimeBrowserSnapshot> =>
      ipcRenderer.invoke(request.animeBrowserGetSnapshot, sessionId),
    ensureBridge: (): Promise<AnimeBrowserBridgeState> =>
      ipcRenderer.invoke(request.animeBrowserEnsureBridge, sessionId),
    updateBridge: (): Promise<AnimeBrowserBridgeState> =>
      ipcRenderer.invoke(request.animeBrowserUpdateBridge),
    selectSource: (sourceId: string): Promise<void> =>
      ipcRenderer.invoke(request.animeBrowserSelectSource, sessionId, sourceId),
    search: (query: string, page?: number): Promise<AnimeBrowserSearchResult> =>
      ipcRenderer.invoke(request.animeBrowserSearch, sessionId, query, page),
    getPopular: (page?: number): Promise<AnimeBrowserSearchResult> =>
      ipcRenderer.invoke(request.animeBrowserGetPopular, sessionId, page),
    getDetails: (animeUrl: string, sourceId?: string): Promise<AnimeBrowserDetails> =>
      ipcRenderer.invoke(request.animeBrowserGetDetails, sessionId, animeUrl, sourceId),
    getEpisodes: (animeUrl: string, sourceId?: string): Promise<AnimeBrowserEpisode[]> =>
      ipcRenderer.invoke(request.animeBrowserGetEpisodes, sessionId, animeUrl, sourceId),
    getWatchState: (
      watchRequest: AnimeBrowserWatchStateRequest,
    ): Promise<AnimeBrowserEpisodeWatchState[]> =>
      ipcRenderer.invoke(request.animeBrowserGetWatchState, watchRequest),
    setWatched: (
      watchedRequest: AnimeBrowserSetWatchedRequest,
    ): Promise<AnimeBrowserEpisodeWatchState[]> =>
      ipcRenderer.invoke(request.animeBrowserSetWatched, watchedRequest),
    playEpisode: (playRequest: AnimeBrowserPlayRequest): Promise<AnimeBrowserPlayResult> =>
      ipcRenderer.invoke(request.animeBrowserPlayEpisode, playRequest),
    queueEpisode: (playRequest: AnimeBrowserPlayRequest): Promise<AnimeBrowserQueueState> =>
      ipcRenderer.invoke(request.animeBrowserQueueEpisode, playRequest),
    dequeueEpisode: (sourceId: string, episodeUrl: string): Promise<AnimeBrowserQueueState> =>
      ipcRenderer.invoke(request.animeBrowserDequeueEpisode, sourceId, episodeUrl),
    clearQueue: (): Promise<AnimeBrowserQueueState> =>
      ipcRenderer.invoke(request.animeBrowserClearQueue),
    getQueue: (): Promise<AnimeBrowserQueueState> =>
      ipcRenderer.invoke(request.animeBrowserGetQueue),
    getPlaybackState: (): Promise<AnimeBrowserPlaybackState | null> =>
      ipcRenderer.invoke(request.animeBrowserGetPlaybackState),
    isPlaying: (): Promise<boolean> => ipcRenderer.invoke(request.animeBrowserIsPlaying),
    getPreferences: (sourceId: string): Promise<SourcePreferenceView[]> =>
      ipcRenderer.invoke(request.animeBrowserGetPreferences, sourceId),
    setPreference: (
      sourceId: string,
      key: string,
      value: string | string[] | boolean,
    ): Promise<SourcePreferenceView[]> =>
      ipcRenderer.invoke(request.animeBrowserSetPreference, sourceId, key, value),
    listAvailableExtensions: (): Promise<AvailableExtensionsResult> =>
      ipcRenderer.invoke(request.animeBrowserListAvailableExtensions),
    installExtension: (pkg: string): Promise<void> =>
      ipcRenderer.invoke(request.animeBrowserInstallExtension, pkg),
    updateAllExtensions: (): Promise<number> =>
      ipcRenderer.invoke(request.animeBrowserUpdateAllExtensions),
    removeExtension: (pkg: string): Promise<void> =>
      ipcRenderer.invoke(request.animeBrowserRemoveExtension, pkg),
    rescanExtensions: (): Promise<void> => ipcRenderer.invoke(request.animeBrowserRescanExtensions),
    addRepo: (url: string): Promise<void> => ipcRenderer.invoke(request.animeBrowserAddRepo, url),
    removeRepo: (url: string): Promise<void> =>
      ipcRenderer.invoke(request.animeBrowserRemoveRepo, url),
    onBridgeState: (listener: (state: AnimeBrowserBridgeState) => void): (() => void) => {
      const handler = (_event: unknown, state: AnimeBrowserBridgeState): void => listener(state);
      ipcRenderer.on(IPC_CHANNELS.event.animeBrowserBridgeState, handler);
      return () => ipcRenderer.removeListener(IPC_CHANNELS.event.animeBrowserBridgeState, handler);
    },
    onSearchUpdate: (listener: (update: AnimeBrowserSearchUpdate) => void): (() => void) => {
      const handler = (_event: unknown, update: AnimeBrowserSearchUpdate): void => listener(update);
      ipcRenderer.on(IPC_CHANNELS.event.animeBrowserSearchUpdate, handler);
      return () => ipcRenderer.removeListener(IPC_CHANNELS.event.animeBrowserSearchUpdate, handler);
    },
    onQueueState: (listener: (state: AnimeBrowserQueueState) => void): (() => void) => {
      const handler = (_event: unknown, state: AnimeBrowserQueueState): void => listener(state);
      ipcRenderer.on(IPC_CHANNELS.event.animeBrowserQueueState, handler);
      return () => ipcRenderer.removeListener(IPC_CHANNELS.event.animeBrowserQueueState, handler);
    },
    onPlaybackState: (
      listener: (state: AnimeBrowserPlaybackState | null) => void,
    ): (() => void) => {
      const handler = (_event: unknown, state: AnimeBrowserPlaybackState | null): void =>
        listener(state);
      ipcRenderer.on(IPC_CHANNELS.event.animeBrowserPlaybackState, handler);
      return () =>
        ipcRenderer.removeListener(IPC_CHANNELS.event.animeBrowserPlaybackState, handler);
    },
  };
}

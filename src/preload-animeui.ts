import { contextBridge, ipcRenderer } from 'electron';
import { IPC_CHANNELS } from './shared/ipc/contracts';
import type {
  AnimeBrowserAPI,
  AnimeBrowserBridgeState,
  AnimeBrowserDetails,
  AnimeBrowserEpisode,
  AnimeBrowserEpisodeWatchState,
  AnimeBrowserWatchStateRequest,
  AnimeBrowserSetWatchedRequest,
  AnimeBrowserPlayRequest,
  AnimeBrowserPlayResult,
  AnimeBrowserQueueState,
  AnimeBrowserSearchResult,
  AnimeBrowserSearchUpdate,
  AnimeBrowserSnapshot,
  AvailableExtensionsResult,
  SourcePreferenceView,
} from './types/anime-browser';

const request = IPC_CHANNELS.request;

const animeBrowserAPI: AnimeBrowserAPI = {
  getSnapshot: (): Promise<AnimeBrowserSnapshot> =>
    ipcRenderer.invoke(request.animeBrowserGetSnapshot),
  ensureBridge: (): Promise<AnimeBrowserBridgeState> =>
    ipcRenderer.invoke(request.animeBrowserEnsureBridge),
  selectSource: (sourceId: string): Promise<void> =>
    ipcRenderer.invoke(request.animeBrowserSelectSource, sourceId),
  search: (query: string, page?: number): Promise<AnimeBrowserSearchResult> =>
    ipcRenderer.invoke(request.animeBrowserSearch, query, page),
  getPopular: (page?: number): Promise<AnimeBrowserSearchResult> =>
    ipcRenderer.invoke(request.animeBrowserGetPopular, page),
  getDetails: (animeUrl: string, sourceId?: string): Promise<AnimeBrowserDetails> =>
    ipcRenderer.invoke(request.animeBrowserGetDetails, animeUrl, sourceId),
  getEpisodes: (animeUrl: string, sourceId?: string): Promise<AnimeBrowserEpisode[]> =>
    ipcRenderer.invoke(request.animeBrowserGetEpisodes, animeUrl, sourceId),
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
  getQueue: (): Promise<AnimeBrowserQueueState> => ipcRenderer.invoke(request.animeBrowserGetQueue),
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
};

contextBridge.exposeInMainWorld('animeBrowserAPI', animeBrowserAPI);

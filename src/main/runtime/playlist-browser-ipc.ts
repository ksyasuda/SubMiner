import type { RegisterIpcRuntimeServicesParams } from '../ipc-runtime';
import {
  appendPlaylistBrowserFileRuntime,
  getPlaylistBrowserSnapshotRuntime,
  movePlaylistBrowserIndexRuntime,
  playPlaylistBrowserIndexRuntime,
  removePlaylistBrowserIndexRuntime,
  type PlaylistBrowserRuntimeDeps,
} from './playlist-browser-runtime';

type PlaylistBrowserMainDeps = Pick<
  RegisterIpcRuntimeServicesParams['mainDeps'],
  | 'getPlaylistBrowserSnapshot'
  | 'appendPlaylistBrowserFile'
  | 'playPlaylistBrowserIndex'
  | 'removePlaylistBrowserIndex'
  | 'movePlaylistBrowserIndex'
>;

export type PlaylistBrowserIpcRuntime = {
  playlistBrowserRuntimeDeps: PlaylistBrowserRuntimeDeps;
  playlistBrowserMainDeps: PlaylistBrowserMainDeps;
};

export function createPlaylistBrowserIpcRuntime(
  getMpvClient: PlaylistBrowserRuntimeDeps['getMpvClient'],
): PlaylistBrowserIpcRuntime {
  const playlistBrowserRuntimeDeps: PlaylistBrowserRuntimeDeps = {
    getMpvClient,
  };

  return {
    playlistBrowserRuntimeDeps,
    playlistBrowserMainDeps: {
      getPlaylistBrowserSnapshot: () => getPlaylistBrowserSnapshotRuntime(playlistBrowserRuntimeDeps),
      appendPlaylistBrowserFile: (filePath: string) =>
        appendPlaylistBrowserFileRuntime(playlistBrowserRuntimeDeps, filePath),
      playPlaylistBrowserIndex: (index: number) =>
        playPlaylistBrowserIndexRuntime(playlistBrowserRuntimeDeps, index),
      removePlaylistBrowserIndex: (index: number) =>
        removePlaylistBrowserIndexRuntime(playlistBrowserRuntimeDeps, index),
      movePlaylistBrowserIndex: (index: number, direction: 1 | -1) =>
        movePlaylistBrowserIndexRuntime(playlistBrowserRuntimeDeps, index, direction),
    },
  };
}

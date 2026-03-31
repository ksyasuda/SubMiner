import fs from 'node:fs';
import path from 'node:path';
import type {
  PlaylistBrowserDirectoryItem,
  PlaylistBrowserMutationResult,
  PlaylistBrowserQueueItem,
  PlaylistBrowserSnapshot,
} from '../../types';
import { isRemoteMediaPath } from '../../jimaku/utils';
import { hasVideoExtension } from '../../shared/video-extensions';
import { sortPlaylistBrowserDirectoryItems } from './playlist-browser-sort';

type PlaylistLike = {
  filename?: unknown;
  title?: unknown;
  id?: unknown;
  current?: unknown;
  playing?: unknown;
};

type MpvPlaylistBrowserClientLike = {
  connected: boolean;
  currentVideoPath?: string | null;
  requestProperty?: (name: string) => Promise<unknown>;
  send: (payload: { command: unknown[]; request_id?: number }) => boolean;
};

export type PlaylistBrowserRuntimeDeps = {
  getMpvClient: () => MpvPlaylistBrowserClientLike | null;
  schedule?: (callback: () => void, delayMs: number) => void;
};

const pendingLocalSubtitleSelectionRearms = new WeakMap<MpvPlaylistBrowserClientLike, number>();

function trimToNull(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

async function readProperty(
  client: MpvPlaylistBrowserClientLike | null,
  name: string,
): Promise<unknown> {
  if (!client?.requestProperty) return null;
  try {
    return await client.requestProperty(name);
  } catch {
    return null;
  }
}

async function resolveCurrentFilePath(
  client: MpvPlaylistBrowserClientLike | null,
): Promise<string | null> {
  const currentVideoPath = trimToNull(client?.currentVideoPath);
  if (currentVideoPath) return currentVideoPath;
  return trimToNull(await readProperty(client, 'path'));
}

function resolveDirectorySnapshot(
  currentFilePath: string | null,
): Pick<PlaylistBrowserSnapshot, 'directoryAvailable' | 'directoryItems' | 'directoryPath' | 'directoryStatus'> {
  if (!currentFilePath) {
    return {
      directoryAvailable: false,
      directoryItems: [],
      directoryPath: null,
      directoryStatus: 'Current media path is unavailable.',
    };
  }

  if (isRemoteMediaPath(currentFilePath)) {
    return {
      directoryAvailable: false,
      directoryItems: [],
      directoryPath: null,
      directoryStatus: 'Directory browser requires a local filesystem video.',
    };
  }

  const resolvedPath = path.resolve(currentFilePath);
  const directoryPath = path.dirname(resolvedPath);
  try {
    const entries = fs.readdirSync(directoryPath, { withFileTypes: true });
    const videoPaths = entries
      .filter((entry) => entry.isFile())
      .map((entry) => entry.name)
      .filter((name) => hasVideoExtension(path.extname(name)))
      .map((name) => path.join(directoryPath, name));

    const directoryItems: PlaylistBrowserDirectoryItem[] = sortPlaylistBrowserDirectoryItems(
      videoPaths,
    ).map((item) => ({
      ...item,
      isCurrentFile: item.path === resolvedPath,
    }));

    return {
      directoryAvailable: true,
      directoryItems,
      directoryPath,
      directoryStatus: directoryPath,
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return {
      directoryAvailable: false,
      directoryItems: [],
      directoryPath,
      directoryStatus: `Could not read parent directory: ${message}`,
    };
  }
}

function normalizePlaylistItems(raw: unknown): PlaylistBrowserQueueItem[] {
  if (!Array.isArray(raw)) return [];
  return raw.map((entry, index) => {
    const item = (entry ?? {}) as PlaylistLike;
    const filename = trimToNull(item.filename) ?? '';
    const title = trimToNull(item.title);
    const normalizedPath =
      filename && !isRemoteMediaPath(filename) ? path.resolve(filename) : trimToNull(filename);
    return {
      index,
      id: typeof item.id === 'number' ? item.id : null,
      filename,
      title,
      displayLabel:
        title ?? (path.basename(filename || '') || filename || `Playlist item ${index + 1}`),
      current: item.current === true,
      playing: item.playing === true,
      path: normalizedPath,
    };
  });
}

function ensureConnectedClient(
  deps: PlaylistBrowserRuntimeDeps,
): MpvPlaylistBrowserClientLike | { ok: false; message: string } {
  const client = deps.getMpvClient();
  if (!client?.connected) {
    return {
      ok: false,
      message: 'MPV is not connected.',
    };
  }
  return client;
}

async function getPlaylistItemsFromClient(
  client: MpvPlaylistBrowserClientLike | null,
): Promise<PlaylistBrowserQueueItem[]> {
  return normalizePlaylistItems(await readProperty(client, 'playlist'));
}

export async function getPlaylistBrowserSnapshotRuntime(
  deps: PlaylistBrowserRuntimeDeps,
): Promise<PlaylistBrowserSnapshot> {
  const client = deps.getMpvClient();
  const currentFilePath = await resolveCurrentFilePath(client);
  const [playlistItems, playingPosValue] = await Promise.all([
    getPlaylistItemsFromClient(client),
    readProperty(client, 'playlist-playing-pos'),
  ]);
  const playingIndex =
    typeof playingPosValue === 'number' && Number.isInteger(playingPosValue)
      ? playingPosValue
      : playlistItems.findIndex((item) => item.current || item.playing);

  return {
    ...resolveDirectorySnapshot(currentFilePath),
    playlistItems,
    playingIndex: playingIndex >= 0 ? playingIndex : null,
    currentFilePath,
  };
}

async function validatePlaylistIndex(
  deps: PlaylistBrowserRuntimeDeps,
  index: number,
): Promise<
  | { ok: false; message: string }
  | { ok: true; client: MpvPlaylistBrowserClientLike; playlistItems: PlaylistBrowserQueueItem[] }
> {
  const client = ensureConnectedClient(deps);
  if ('ok' in client) {
    return client;
  }
  const playlistItems = await getPlaylistItemsFromClient(client);
  if (!Number.isInteger(index) || index < 0 || index >= playlistItems.length) {
    return {
      ok: false,
      message: 'Playlist item not found.',
    };
  }
  return {
    ok: true,
    client,
    playlistItems,
  };
}

async function buildMutationResult(
  message: string,
  deps: PlaylistBrowserRuntimeDeps,
): Promise<PlaylistBrowserMutationResult> {
  return {
    ok: true,
    message,
    snapshot: await getPlaylistBrowserSnapshotRuntime(deps),
  };
}

function rearmLocalSubtitleSelection(client: MpvPlaylistBrowserClientLike): void {
  client.send({ command: ['set_property', 'sid', 'auto'] });
  client.send({ command: ['set_property', 'secondary-sid', 'auto'] });
}

function prepareLocalSubtitleAutoload(client: MpvPlaylistBrowserClientLike): void {
  client.send({ command: ['set_property', 'sub-auto', 'fuzzy'] });
}

function isLocalPlaylistItem(
  item: PlaylistBrowserQueueItem | null | undefined,
): item is PlaylistBrowserQueueItem & { path: string } {
  return Boolean(item?.path && !isRemoteMediaPath(item.path));
}

function scheduleLocalSubtitleSelectionRearm(
  deps: PlaylistBrowserRuntimeDeps,
  client: MpvPlaylistBrowserClientLike,
  expectedPath: string,
): void {
  const nextToken = (pendingLocalSubtitleSelectionRearms.get(client) ?? 0) + 1;
  pendingLocalSubtitleSelectionRearms.set(client, nextToken);
  (deps.schedule ?? setTimeout)(() => {
    if (pendingLocalSubtitleSelectionRearms.get(client) !== nextToken) return;
    pendingLocalSubtitleSelectionRearms.delete(client);
    const currentPath = trimToNull(client.currentVideoPath);
    if (currentPath && path.resolve(currentPath) !== expectedPath) {
      return;
    }
    rearmLocalSubtitleSelection(client);
  }, 400);
}

export async function appendPlaylistBrowserFileRuntime(
  deps: PlaylistBrowserRuntimeDeps,
  filePath: string,
): Promise<PlaylistBrowserMutationResult> {
  const client = ensureConnectedClient(deps);
  if ('ok' in client) {
    return client;
  }
  const resolvedPath = path.resolve(filePath);
  let stats: fs.Stats;
  try {
    stats = fs.statSync(resolvedPath);
  } catch {
    return {
      ok: false,
      message: 'Playlist browser file is not readable.',
    };
  }
  if (!stats.isFile()) {
    return {
      ok: false,
      message: 'Playlist browser file is not readable.',
    };
  }

  client.send({ command: ['loadfile', resolvedPath, 'append'] });
  return buildMutationResult(`Queued ${path.basename(resolvedPath)}`, deps);
}

export async function playPlaylistBrowserIndexRuntime(
  deps: PlaylistBrowserRuntimeDeps,
  index: number,
): Promise<PlaylistBrowserMutationResult> {
  const result = await validatePlaylistIndex(deps, index);
  if (!result.ok) {
    return result;
  }

  const targetItem = result.playlistItems[index] ?? null;
  if (isLocalPlaylistItem(targetItem)) {
    prepareLocalSubtitleAutoload(result.client);
  }
  result.client.send({ command: ['playlist-play-index', index] });
  if (isLocalPlaylistItem(targetItem)) {
    scheduleLocalSubtitleSelectionRearm(deps, result.client, path.resolve(targetItem.path));
  }
  return buildMutationResult(`Playing playlist item ${index + 1}`, deps);
}

export async function removePlaylistBrowserIndexRuntime(
  deps: PlaylistBrowserRuntimeDeps,
  index: number,
): Promise<PlaylistBrowserMutationResult> {
  const result = await validatePlaylistIndex(deps, index);
  if (!result.ok) {
    return result;
  }

  result.client.send({ command: ['playlist-remove', index] });
  return buildMutationResult(`Removed playlist item ${index + 1}`, deps);
}

export async function movePlaylistBrowserIndexRuntime(
  deps: PlaylistBrowserRuntimeDeps,
  index: number,
  direction: -1 | 1,
): Promise<PlaylistBrowserMutationResult> {
  const result = await validatePlaylistIndex(deps, index);
  if (!result.ok) {
    return result;
  }

  const targetIndex = index + direction;
  if (targetIndex < 0) {
    return {
      ok: false,
      message: 'Playlist item is already at the top.',
    };
  }
  if (targetIndex >= result.playlistItems.length) {
    return {
      ok: false,
      message: 'Playlist item is already at the bottom.',
    };
  }

  result.client.send({ command: ['playlist-move', index, targetIndex] });
  return buildMutationResult(`Moved playlist item ${index + 1}`, deps);
}

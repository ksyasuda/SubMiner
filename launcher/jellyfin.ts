import path from 'node:path';
import fs from 'node:fs';
import { spawnSync } from 'node:child_process';
import type {
  Args,
  JellyfinSessionConfig,
  JellyfinLibraryEntry,
  JellyfinItemEntry,
  JellyfinGroupEntry,
} from './types.js';
import { log, fail } from './log.js';
import { commandExists, resolvePathMaybe } from './util.js';
import {
  pickLibrary,
  pickItem,
  pickGroup,
  promptOptionalJellyfinSearch,
  findRofiTheme,
} from './picker.js';
import { loadLauncherJellyfinConfig } from './config.js';
import {
  runAppCommandWithInheritLogged,
  launchMpvIdleDetached,
  waitForUnixSocketReady,
} from './mpv.js';

export function sanitizeServerUrl(value: string): string {
  return value.trim().replace(/\/+$/, '');
}

export async function jellyfinApiRequest<T>(
  session: JellyfinSessionConfig,
  requestPath: string,
): Promise<T> {
  const url = `${session.serverUrl}${requestPath}`;
  const response = await fetch(url, {
    headers: {
      'X-Emby-Token': session.accessToken,
      Authorization: `MediaBrowser Token="${session.accessToken}"`,
    },
  });
  if (response.status === 401 || response.status === 403) {
    fail('Jellyfin token invalid/expired. Run --jellyfin-login or --jellyfin.');
  }
  if (!response.ok) {
    fail(`Jellyfin API failed: ${response.status} ${response.statusText}`);
  }
  return (await response.json()) as T;
}

function itemPreviewUrl(session: JellyfinSessionConfig, id: string): string {
  return `${session.serverUrl}/Items/${id}/Images/Primary?maxHeight=720&quality=85&api_key=${encodeURIComponent(session.accessToken)}`;
}

function jellyfinIconCacheDir(session: JellyfinSessionConfig): string {
  const serverKey = session.serverUrl.replace(/[^a-zA-Z0-9]+/g, '_').slice(0, 96);
  const userKey = session.userId.replace(/[^a-zA-Z0-9]+/g, '_').slice(0, 96);
  const baseDir = session.iconCacheDir
    ? resolvePathMaybe(session.iconCacheDir)
    : path.join('/tmp', 'subminer-jellyfin-icons');
  return path.join(baseDir, serverKey, userKey);
}

function jellyfinIconPath(session: JellyfinSessionConfig, id: string): string {
  const safeId = id.replace(/[^a-zA-Z0-9._-]+/g, '_');
  return path.join(jellyfinIconCacheDir(session), `${safeId}.jpg`);
}

function ensureJellyfinIcon(session: JellyfinSessionConfig, id: string): string | null {
  if (!session.pullPictures || !id || !commandExists('curl')) return null;
  const iconPath = jellyfinIconPath(session, id);
  try {
    if (fs.existsSync(iconPath) && fs.statSync(iconPath).size > 0) {
      return iconPath;
    }
  } catch {
    // continue to download
  }

  try {
    fs.mkdirSync(path.dirname(iconPath), { recursive: true });
  } catch {
    return null;
  }

  const result = spawnSync('curl', ['-fsSL', '-o', iconPath, itemPreviewUrl(session, id)], {
    stdio: 'ignore',
  });
  if (result.error || result.status !== 0) return null;

  try {
    if (fs.existsSync(iconPath) && fs.statSync(iconPath).size > 0) {
      return iconPath;
    }
  } catch {
    return null;
  }
  return null;
}

export function formatJellyfinItemDisplay(item: Record<string, unknown>): string {
  const type = typeof item.Type === 'string' ? item.Type : 'Item';
  const name = typeof item.Name === 'string' ? item.Name : 'Untitled';
  if (type === 'Episode') {
    const series = typeof item.SeriesName === 'string' ? item.SeriesName : '';
    const season =
      typeof item.ParentIndexNumber === 'number'
        ? String(item.ParentIndexNumber).padStart(2, '0')
        : '00';
    const episode =
      typeof item.IndexNumber === 'number' ? String(item.IndexNumber).padStart(2, '0') : '00';
    return `${series} S${season}E${episode} ${name}`.trim();
  }
  return `${name} (${type})`;
}

export async function resolveJellyfinSelection(
  args: Args,
  session: JellyfinSessionConfig,
  themePath: string | null = null,
): Promise<string> {
  const asNumberOrNull = (value: unknown): number | null => {
    if (typeof value !== 'number' || !Number.isFinite(value)) return null;
    return value;
  };
  const compareByName = (left: string, right: string): number =>
    left.localeCompare(right, undefined, { sensitivity: 'base', numeric: true });
  const sortEntries = (
    entries: Array<{
      type: string;
      name: string;
      parentIndex: number | null;
      index: number | null;
      display: string;
    }>,
  ) =>
    entries.sort((left, right) => {
      if (left.type === 'Episode' && right.type === 'Episode') {
        const leftSeason = left.parentIndex ?? Number.MAX_SAFE_INTEGER;
        const rightSeason = right.parentIndex ?? Number.MAX_SAFE_INTEGER;
        if (leftSeason !== rightSeason) return leftSeason - rightSeason;
        const leftEpisode = left.index ?? Number.MAX_SAFE_INTEGER;
        const rightEpisode = right.index ?? Number.MAX_SAFE_INTEGER;
        if (leftEpisode !== rightEpisode) return leftEpisode - rightEpisode;
      }
      if (left.type !== right.type) {
        const leftEpisodeLike = left.type === 'Episode';
        const rightEpisodeLike = right.type === 'Episode';
        if (leftEpisodeLike && !rightEpisodeLike) return -1;
        if (!leftEpisodeLike && rightEpisodeLike) return 1;
      }
      return compareByName(left.display, right.display);
    });

  const libsPayload = await jellyfinApiRequest<{ Items?: Array<Record<string, unknown>> }>(
    session,
    `/Users/${session.userId}/Views`,
  );
  const libraries: JellyfinLibraryEntry[] = (libsPayload.Items || [])
    .map((item) => ({
      id: typeof item.Id === 'string' ? item.Id : '',
      name: typeof item.Name === 'string' ? item.Name : 'Untitled',
      kind:
        typeof item.CollectionType === 'string'
          ? item.CollectionType
          : typeof item.Type === 'string'
            ? item.Type
            : 'unknown',
    }))
    .filter((item) => item.id.length > 0);

  let libraryId = session.defaultLibraryId;
  if (!libraryId) {
    libraryId = pickLibrary(session, libraries, args.useRofi, ensureJellyfinIcon, '', themePath);
    if (!libraryId) fail('No Jellyfin library selected.');
  }
  const searchTerm = await promptOptionalJellyfinSearch(args.useRofi, themePath);

  const fetchItemsPaged = async (parentId: string) => {
    const out: Array<Record<string, unknown>> = [];
    let startIndex = 0;
    while (true) {
      const payload = await jellyfinApiRequest<{
        Items?: Array<Record<string, unknown>>;
        TotalRecordCount?: number;
      }>(
        session,
        `/Users/${session.userId}/Items?ParentId=${encodeURIComponent(parentId)}&Recursive=false&SortBy=SortName&SortOrder=Ascending&Limit=500&StartIndex=${startIndex}`,
      );
      const page = payload.Items || [];
      if (page.length === 0) break;
      out.push(...page);
      startIndex += page.length;
      const total = typeof payload.TotalRecordCount === 'number' ? payload.TotalRecordCount : null;
      if (total !== null && startIndex >= total) break;
      if (page.length < 500) break;
    }
    return out;
  };

  const topLevelEntries = await fetchItemsPaged(libraryId);
  const groups: JellyfinGroupEntry[] = topLevelEntries
    .filter((item) => {
      const type = typeof item.Type === 'string' ? item.Type : '';
      return (
        type === 'Series' || type === 'Folder' || type === 'CollectionFolder' || type === 'Season'
      );
    })
    .map((item) => {
      const type = typeof item.Type === 'string' ? item.Type : 'Folder';
      const name = typeof item.Name === 'string' ? item.Name : 'Untitled';
      return {
        id: typeof item.Id === 'string' ? item.Id : '',
        name,
        type,
        display: `${name} (${type})`,
      };
    })
    .filter((entry) => entry.id.length > 0);

  let contentParentId = libraryId;
  let contentRecursive = true;
  const selectedGroupId = pickGroup(
    session,
    groups,
    args.useRofi,
    ensureJellyfinIcon,
    searchTerm,
    themePath,
  );
  if (selectedGroupId) {
    contentParentId = selectedGroupId;
    const nextLevelEntries = await fetchItemsPaged(selectedGroupId);
    const seasons: JellyfinGroupEntry[] = nextLevelEntries
      .filter((item) => {
        const type = typeof item.Type === 'string' ? item.Type : '';
        return type === 'Season' || type === 'Folder';
      })
      .map((item) => {
        const type = typeof item.Type === 'string' ? item.Type : 'Season';
        const name = typeof item.Name === 'string' ? item.Name : 'Untitled';
        return {
          id: typeof item.Id === 'string' ? item.Id : '',
          name,
          type,
          display: `${name} (${type})`,
        };
      })
      .filter((entry) => entry.id.length > 0);
    if (seasons.length > 0) {
      const seasonsById = new Map(seasons.map((entry) => [entry.id, entry]));
      const selectedSeasonId = pickGroup(
        session,
        seasons,
        args.useRofi,
        ensureJellyfinIcon,
        '',
        themePath,
      );
      if (!selectedSeasonId) fail('No Jellyfin season selected.');
      contentParentId = selectedSeasonId;
      const selectedSeason = seasonsById.get(selectedSeasonId);
      if (selectedSeason?.type === 'Season') {
        contentRecursive = false;
      }
    }
  }

  const fetchPage = async (startIndex: number) =>
    jellyfinApiRequest<{
      Items?: Array<Record<string, unknown>>;
      TotalRecordCount?: number;
    }>(
      session,
      `/Users/${session.userId}/Items?ParentId=${encodeURIComponent(contentParentId)}&Recursive=${contentRecursive ? 'true' : 'false'}&SortBy=SortName&SortOrder=Ascending&Limit=500&StartIndex=${startIndex}`,
    );

  const allEntries: Array<Record<string, unknown>> = [];
  let startIndex = 0;
  while (true) {
    const payload = await fetchPage(startIndex);
    const page = payload.Items || [];
    if (page.length === 0) break;
    allEntries.push(...page);
    startIndex += page.length;
    const total = typeof payload.TotalRecordCount === 'number' ? payload.TotalRecordCount : null;
    if (total !== null && startIndex >= total) break;
    if (page.length < 500) break;
  }

  let items: JellyfinItemEntry[] = sortEntries(
    allEntries
      .filter((item) => {
        const type = typeof item.Type === 'string' ? item.Type : '';
        return type === 'Movie' || type === 'Episode' || type === 'Audio';
      })
      .map((item) => ({
        id: typeof item.Id === 'string' ? item.Id : '',
        name: typeof item.Name === 'string' ? item.Name : '',
        type: typeof item.Type === 'string' ? item.Type : 'Item',
        parentIndex: asNumberOrNull(item.ParentIndexNumber),
        index: asNumberOrNull(item.IndexNumber),
        display: formatJellyfinItemDisplay(item),
      }))
      .filter((item) => item.id.length > 0),
  ).map(({ id, name, type, display }) => ({
    id,
    name,
    type,
    display,
  }));

  if (items.length === 0) {
    items = sortEntries(
      allEntries
        .filter((item) => {
          const type = typeof item.Type === 'string' ? item.Type : '';
          if (type === 'Folder' || type === 'CollectionFolder') return false;
          const mediaType = typeof item.MediaType === 'string' ? item.MediaType.toLowerCase() : '';
          if (mediaType === 'video' || mediaType === 'audio') return true;
          return (
            type === 'Movie' ||
            type === 'Episode' ||
            type === 'Audio' ||
            type === 'Video' ||
            type === 'MusicVideo'
          );
        })
        .map((item) => ({
          id: typeof item.Id === 'string' ? item.Id : '',
          name: typeof item.Name === 'string' ? item.Name : '',
          type: typeof item.Type === 'string' ? item.Type : 'Item',
          parentIndex: asNumberOrNull(item.ParentIndexNumber),
          index: asNumberOrNull(item.IndexNumber),
          display: formatJellyfinItemDisplay(item),
        }))
        .filter((item) => item.id.length > 0),
    ).map(({ id, name, type, display }) => ({
      id,
      name,
      type,
      display,
    }));
  }

  const itemId = pickItem(session, items, args.useRofi, ensureJellyfinIcon, '', themePath);
  if (!itemId) fail('No Jellyfin item selected.');
  return itemId;
}

export async function runJellyfinPlayMenu(
  appPath: string,
  args: Args,
  scriptPath: string,
  mpvSocketPath: string,
): Promise<never> {
  const config = loadLauncherJellyfinConfig();
  const session: JellyfinSessionConfig = {
    serverUrl: sanitizeServerUrl(args.jellyfinServer || config.serverUrl || ''),
    accessToken: config.accessToken || '',
    userId: config.userId || '',
    defaultLibraryId: config.defaultLibraryId || '',
    pullPictures: config.pullPictures === true,
    iconCacheDir: config.iconCacheDir || '',
  };

  if (!session.serverUrl || !session.accessToken || !session.userId) {
    fail(
      'Missing Jellyfin session config. Run `subminer --jellyfin` or `subminer --jellyfin-login` first.',
    );
  }

  const rofiTheme = args.useRofi ? findRofiTheme(scriptPath) : null;
  if (args.useRofi && !rofiTheme) {
    log('warn', args.logLevel, 'Rofi theme not found for Jellyfin picker; using rofi defaults.');
  }

  const itemId = await resolveJellyfinSelection(args, session, rofiTheme);
  log('debug', args.logLevel, `Jellyfin selection resolved: itemId=${itemId}`);
  log('debug', args.logLevel, `Ensuring MPV IPC socket is ready: ${mpvSocketPath}`);
  let mpvReady = false;
  if (fs.existsSync(mpvSocketPath)) {
    mpvReady = await waitForUnixSocketReady(mpvSocketPath, 250);
  }
  if (!mpvReady) {
    await launchMpvIdleDetached(mpvSocketPath, appPath, args);
    mpvReady = await waitForUnixSocketReady(mpvSocketPath, 8000);
  }
  log('debug', args.logLevel, `MPV socket ready check result: ${mpvReady ? 'ready' : 'not ready'}`);
  if (!mpvReady) {
    fail(`MPV IPC socket not ready: ${mpvSocketPath}`);
  }
  const forwarded = ['--start', '--jellyfin-play', '--jellyfin-item-id', itemId];
  if (args.logLevel !== 'info') forwarded.push('--log-level', args.logLevel);
  runAppCommandWithInheritLogged(appPath, forwarded, args.logLevel, 'jellyfin-play');
}

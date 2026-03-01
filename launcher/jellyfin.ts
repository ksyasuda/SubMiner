import path from 'node:path';
import fs from 'node:fs';
import os from 'node:os';
import { spawnSync } from 'node:child_process';
import type {
  Args,
  JellyfinSessionConfig,
  JellyfinLibraryEntry,
  JellyfinItemEntry,
  JellyfinGroupEntry,
} from './types.js';
import { log, fail, getMpvLogPath } from './log.js';
import { commandExists, resolvePathMaybe, sleep } from './util.js';
import {
  pickLibrary,
  pickItem,
  pickGroup,
  promptOptionalJellyfinSearch,
  findRofiTheme,
} from './picker.js';
import { loadLauncherJellyfinConfig } from './config.js';
import { resolveLauncherMainConfigPath } from './config/shared-config-reader.js';
import {
  runAppCommandWithInheritLogged,
  runAppCommandCaptureOutput,
  launchAppStartDetached,
  launchMpvIdleDetached,
  waitForUnixSocketReady,
} from './mpv.js';

const ANSI_ESCAPE_PATTERN = /\u001b\[[0-9;]*m/g;

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

function stripAnsi(value: string): string {
  return value.replace(ANSI_ESCAPE_PATTERN, '');
}

function parseNamedJellyfinRecord(payload: string): {
  name: string;
  id: string;
  type: string;
} | null {
  const typeClose = payload.lastIndexOf(')');
  if (typeClose !== payload.length - 1) return null;

  const typeOpen = payload.lastIndexOf(' (');
  if (typeOpen <= 0 || typeOpen >= typeClose) return null;

  const idClose = payload.lastIndexOf(']', typeOpen);
  if (idClose <= 0) return null;

  const idOpen = payload.lastIndexOf(' [', idClose);
  if (idOpen <= 0 || idOpen >= idClose) return null;

  const name = payload.slice(0, idOpen).trim();
  const id = payload.slice(idOpen + 2, idClose).trim();
  const type = payload.slice(typeOpen + 2, typeClose).trim();
  if (!name || !id || !type) return null;

  return { name, id, type };
}

export function parseJellyfinLibrariesFromAppOutput(output: string): JellyfinLibraryEntry[] {
  const libraries: JellyfinLibraryEntry[] = [];
  const seenIds = new Set<string>();

  for (const rawLine of output.split(/\r?\n/)) {
    const line = stripAnsi(rawLine);
    const markerIndex = line.indexOf('Jellyfin library:');
    if (markerIndex < 0) continue;
    const payload = line.slice(markerIndex + 'Jellyfin library:'.length).trim();
    const parsed = parseNamedJellyfinRecord(payload);
    if (!parsed || seenIds.has(parsed.id)) continue;
    seenIds.add(parsed.id);
    libraries.push({
      id: parsed.id,
      name: parsed.name,
      kind: parsed.type,
    });
  }

  return libraries;
}

export function parseJellyfinItemsFromAppOutput(output: string): JellyfinItemEntry[] {
  const items: JellyfinItemEntry[] = [];
  const seenIds = new Set<string>();

  for (const rawLine of output.split(/\r?\n/)) {
    const line = stripAnsi(rawLine);
    const markerIndex = line.indexOf('Jellyfin item:');
    if (markerIndex < 0) continue;
    const payload = line.slice(markerIndex + 'Jellyfin item:'.length).trim();
    const parsed = parseNamedJellyfinRecord(payload);
    if (!parsed || seenIds.has(parsed.id)) continue;
    seenIds.add(parsed.id);
    items.push({
      id: parsed.id,
      name: parsed.name,
      type: parsed.type,
      display: parsed.name,
    });
  }

  return items;
}

export function parseJellyfinErrorFromAppOutput(output: string): string {
  const lines = output.split(/\r?\n/).map((line) => stripAnsi(line).trim());
  for (let i = lines.length - 1; i >= 0; i -= 1) {
    const line = lines[i];
    if (!line) continue;

    const bracketedErrorIndex = line.indexOf('[ERROR]');
    if (bracketedErrorIndex >= 0) {
      const message = line.slice(bracketedErrorIndex + '[ERROR]'.length).trim();
      if (message.length > 0) return message;
    }

    const mainErrorIndex = line.indexOf(' - ERROR - ');
    if (mainErrorIndex >= 0) {
      const message = line.slice(mainErrorIndex + ' - ERROR - '.length).trim();
      if (message.length > 0) return message;
    }

    if (line.includes('Missing Jellyfin session')) {
      return 'Missing Jellyfin session. Run `subminer jellyfin -l` to log in again.';
    }
  }
  return '';
}

type JellyfinPreviewAuthResponse = {
  serverUrl: string;
  accessToken: string;
  userId: string;
};

export function parseJellyfinPreviewAuthResponse(raw: string): JellyfinPreviewAuthResponse | null {
  if (!raw || raw.trim().length === 0) return null;
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return null;
  }
  if (!parsed || typeof parsed !== 'object') return null;

  const candidate = parsed as Record<string, unknown>;
  const serverUrl = sanitizeServerUrl(
    typeof candidate.serverUrl === 'string' ? candidate.serverUrl : '',
  );
  const accessToken =
    typeof candidate.accessToken === 'string' ? candidate.accessToken.trim() : '';
  const userId = typeof candidate.userId === 'string' ? candidate.userId.trim() : '';
  if (!serverUrl || !accessToken) return null;

  return {
    serverUrl,
    accessToken,
    userId,
  };
}

export function shouldRetryWithStartForNoRunningInstance(errorMessage: string): boolean {
  return errorMessage.includes('No running instance. Use --start to launch the app.');
}

export function deriveJellyfinTokenStorePath(configPath: string): string {
  return path.join(path.dirname(configPath), 'jellyfin-token-store.json');
}

export function hasStoredJellyfinSession(
  configPath: string,
  exists: (candidate: string) => boolean = fs.existsSync,
): boolean {
  return exists(deriveJellyfinTokenStorePath(configPath));
}

export function readUtf8FileAppendedSince(logPath: string, offsetBytes: number): string {
  try {
    const buffer = fs.readFileSync(logPath);
    if (buffer.length === 0) return '';
    const normalizedOffset =
      Number.isFinite(offsetBytes) && offsetBytes >= 0
        ? Math.floor(offsetBytes)
        : 0;
    const startOffset = normalizedOffset > buffer.length ? 0 : normalizedOffset;
    return buffer.subarray(startOffset).toString('utf8');
  } catch {
    return '';
  }
}

export function parseEpisodePathFromDisplay(
  display: string,
): { seriesName: string; seasonNumber: number } | null {
  const normalized = display.trim().replace(/\s+/g, ' ');
  const match = normalized.match(/^(.*?)\s+S(\d{1,2})E\d{1,3}\b/i);
  if (!match) return null;
  const seriesName = match[1].trim();
  const seasonNumber = Number.parseInt(match[2], 10);
  if (!seriesName || !Number.isFinite(seasonNumber) || seasonNumber < 0) return null;
  return { seriesName, seasonNumber };
}

function normalizeJellyfinType(type: string): string {
  return type.trim().toLowerCase();
}

export function isJellyfinPlayableType(type: string): boolean {
  const normalizedType = normalizeJellyfinType(type);
  return (
    normalizedType === 'movie' ||
    normalizedType === 'episode' ||
    normalizedType === 'audio' ||
    normalizedType === 'video' ||
    normalizedType === 'musicvideo'
  );
}

export function isJellyfinContainerType(type: string): boolean {
  const normalizedType = normalizeJellyfinType(type);
  return (
    normalizedType === 'series' ||
    normalizedType === 'season' ||
    normalizedType === 'folder' ||
    normalizedType === 'collectionfolder'
  );
}

function isJellyfinRootSearchType(type: string): boolean {
  const normalizedType = normalizeJellyfinType(type);
  return (
    isJellyfinContainerType(normalizedType) ||
    normalizedType === 'movie' ||
    normalizedType === 'video' ||
    normalizedType === 'musicvideo'
  );
}

export function buildRootSearchGroups(items: JellyfinItemEntry[]): JellyfinGroupEntry[] {
  const seenIds = new Set<string>();
  const groups: JellyfinGroupEntry[] = [];
  for (const item of items) {
    if (!item.id || seenIds.has(item.id) || !isJellyfinRootSearchType(item.type)) continue;
    seenIds.add(item.id);
    groups.push({
      id: item.id,
      name: item.name,
      type: item.type,
      display: `${item.name} (${item.type})`,
    });
  }
  return groups;
}

async function runAppJellyfinListCommand(
  appPath: string,
  args: Args,
  appArgs: string[],
  label: string,
): Promise<string> {
  const attempt = await runAppJellyfinCommand(appPath, args, appArgs, label);
  if (attempt.status !== 0) {
    const message = attempt.output.trim();
    fail(message || `${label} failed.`);
  }
  if (attempt.error) {
    fail(attempt.error);
  }
  return attempt.output;
}

async function runAppJellyfinCommand(
  appPath: string,
  args: Args,
  appArgs: string[],
  label: string,
): Promise<{ status: number; output: string; error: string; logOffset: number }> {
  const forwardedBase = [...appArgs];
  const serverOverride = sanitizeServerUrl(args.jellyfinServer || '');
  if (serverOverride) {
    forwardedBase.push('--jellyfin-server', serverOverride);
  }
  if (args.passwordStore) {
    forwardedBase.push('--password-store', args.passwordStore);
  }

  const readLogAppendedSince = (offset: number): string => {
    const logPath = getMpvLogPath();
    return readUtf8FileAppendedSince(logPath, offset);
  };

  const hasCommandSignal = (output: string): boolean => {
    if (label === 'jellyfin-libraries') {
      return output.includes('Jellyfin library:') || output.includes('No Jellyfin libraries found.');
    }
    if (label === 'jellyfin-items') {
      return (
        output.includes('Jellyfin item:') ||
        output.includes('No Jellyfin items found for the selected library/search.')
      );
    }
    if (label === 'jellyfin-preview-auth') {
      return output.includes('Jellyfin preview auth written.');
    }
    return output.trim().length > 0;
  };

  const runOnce = (): { status: number; output: string; error: string; logOffset: number } => {
    const forwarded = [...forwardedBase];
    const logPath = getMpvLogPath();
    let logOffset = 0;
    try {
      if (fs.existsSync(logPath)) {
        logOffset = fs.statSync(logPath).size;
      }
    } catch {
      logOffset = 0;
    }
    log('debug', args.logLevel, `${label}: launching app with args: ${forwarded.join(' ')}`);
    const result = runAppCommandCaptureOutput(appPath, forwarded);
    log('debug', args.logLevel, `${label}: app command exited with status ${result.status}`);
    let output = `${result.stdout || ''}\n${result.stderr || ''}\n${readLogAppendedSince(logOffset)}`;
    let error = parseJellyfinErrorFromAppOutput(output);

    return { status: result.status, output, error, logOffset };
  };

  let retriedAfterStart = false;
  let attempt = runOnce();
  if (shouldRetryWithStartForNoRunningInstance(attempt.error)) {
    log('debug', args.logLevel, `${label}: starting app detached, then retrying command`);
    launchAppStartDetached(appPath, args.logLevel);
    await sleep(1000);
    retriedAfterStart = true;
    attempt = runOnce();
  }

  if (attempt.status === 0 && !attempt.error && !hasCommandSignal(attempt.output)) {
    // When app is already running, command handling happens in the primary process and log
    // lines can land slightly after the helper process exits.
    const settleWindowMs = (() => {
      if (label === 'jellyfin-items') {
        return retriedAfterStart ? 45000 : 30000;
      }
      return retriedAfterStart ? 12000 : 4000;
    })();
    const settleDeadline = Date.now() + settleWindowMs;
    const settleOffset = attempt.logOffset;
    while (Date.now() < settleDeadline) {
      await sleep(100);
      const settledOutput = readLogAppendedSince(settleOffset);
      if (!settledOutput.trim()) {
        continue;
      }
      attempt.output = `${attempt.output}\n${settledOutput}`;
      attempt.error = parseJellyfinErrorFromAppOutput(attempt.output);
      if (attempt.error || hasCommandSignal(attempt.output)) {
        break;
      }
    }
  }

  return attempt;
}

async function requestJellyfinPreviewAuthFromApp(
  appPath: string,
  args: Args,
): Promise<JellyfinPreviewAuthResponse | null> {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-jf-preview-auth-'));
  const responsePath = path.join(tmpDir, 'response.json');
  try {
    const attempt = await runAppJellyfinCommand(
      appPath,
      args,
      ['--jellyfin-preview-auth', `--jellyfin-response-path=${responsePath}`],
      'jellyfin-preview-auth',
    );
    if (attempt.status !== 0 || attempt.error) {
      return null;
    }

    const deadline = Date.now() + 4000;
    while (Date.now() < deadline) {
      try {
        if (fs.existsSync(responsePath)) {
          const raw = fs.readFileSync(responsePath, 'utf8');
          const parsed = parseJellyfinPreviewAuthResponse(raw);
          if (parsed) {
            return parsed;
          }
        }
      } catch {
        // retry until timeout
      }
      await sleep(100);
    }
    return null;
  } finally {
    try {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    } catch {
      // ignore cleanup failures
    }
  }
}

async function resolveJellyfinSelectionViaApp(
  appPath: string,
  args: Args,
  session: JellyfinSessionConfig,
  themePath: string | null = null,
): Promise<string> {
  const listLibrariesOutput = await runAppJellyfinListCommand(
    appPath,
    args,
    ['--jellyfin-libraries'],
    'jellyfin-libraries',
  );
  const libraries = parseJellyfinLibrariesFromAppOutput(listLibrariesOutput);
  if (libraries.length === 0) {
    fail('No Jellyfin libraries found.');
  }

  const iconlessSession: JellyfinSessionConfig = {
    ...session,
    userId: session.userId || 'launcher',
  };
  const noIcon = (): string | null => null;
  const hasPreviewSession = Boolean(iconlessSession.serverUrl && iconlessSession.accessToken);
  const pickerSession: JellyfinSessionConfig = {
    ...iconlessSession,
    pullPictures: hasPreviewSession && iconlessSession.pullPictures === true,
  };
  const ensureIconForPicker = hasPreviewSession ? ensureJellyfinIcon : noIcon;
  if (!hasPreviewSession) {
    log(
      'debug',
      args.logLevel,
      'Jellyfin picker image previews disabled (no launcher-accessible Jellyfin token).',
    );
  }

  const configuredDefaultLibraryId = session.defaultLibraryId;
  const hasConfiguredDefault = libraries.some((library) => library.id === configuredDefaultLibraryId);
  let libraryId = hasConfiguredDefault ? configuredDefaultLibraryId : '';
  if (!libraryId) {
    libraryId = pickLibrary(
      pickerSession,
      libraries,
      args.useRofi,
      ensureIconForPicker,
      '',
      themePath,
    );
    if (!libraryId) fail('No Jellyfin library selected.');
  }

  const searchTerm = await promptOptionalJellyfinSearch(args.useRofi, themePath);
  const normalizedSearch = searchTerm.trim();
  const searchLimit = 400;
  const browseLimit = 2500;
  const rootIncludeItemTypes = 'Series,Season,Folder,CollectionFolder,Movie,Video,MusicVideo';
  const directoryIncludeItemTypes =
    'Series,Season,Folder,CollectionFolder,Movie,Episode,Audio,Video,MusicVideo';
  const recursivePlayableIncludeItemTypes = 'Movie,Episode,Audio,Video,MusicVideo';
  const listItemsViaApp = async (
    parentId: string,
    options: {
      search?: string;
      limit: number;
      recursive?: boolean;
      includeItemTypes?: string;
    },
  ): Promise<JellyfinItemEntry[]> => {
    const itemArgs = [
      '--jellyfin-items',
      `--jellyfin-library-id=${parentId}`,
      `--jellyfin-limit=${Math.max(1, options.limit)}`,
    ];
    const normalized = (options.search || '').trim();
    if (normalized.length > 0) {
      itemArgs.push(`--jellyfin-search=${normalized}`);
    }
    if (typeof options.recursive === 'boolean') {
      itemArgs.push(`--jellyfin-recursive=${options.recursive ? 'true' : 'false'}`);
    }
    const includeItemTypes = options.includeItemTypes?.trim();
    if (includeItemTypes) {
      itemArgs.push(`--jellyfin-include-item-types=${includeItemTypes}`);
    }
    const output = await runAppJellyfinListCommand(appPath, args, itemArgs, 'jellyfin-items');
    return parseJellyfinItemsFromAppOutput(output);
  };

  let rootItems =
    normalizedSearch.length > 0
      ? await listItemsViaApp(libraryId, {
          search: normalizedSearch,
          limit: searchLimit,
          recursive: true,
          includeItemTypes: rootIncludeItemTypes,
        })
      : await listItemsViaApp(libraryId, {
          limit: browseLimit,
          recursive: false,
          includeItemTypes: rootIncludeItemTypes,
        });
  if (normalizedSearch.length > 0 && rootItems.length === 0) {
    // Compatibility fallback for older app binaries that may ignore custom search include types.
    log(
      'debug',
      args.logLevel,
      `jellyfin-items: no direct search hits for "${normalizedSearch}", falling back to unfiltered library query`,
    );
    rootItems = await listItemsViaApp(libraryId, {
      limit: browseLimit,
      recursive: false,
      includeItemTypes: rootIncludeItemTypes,
    });
  }
  const rootGroups = buildRootSearchGroups(rootItems);
  if (rootGroups.length === 0) {
    fail('No Jellyfin shows or movies found.');
  }

  const rootById = new Map(rootGroups.map((group) => [group.id, group]));
  const selectedRootId = pickGroup(
    pickerSession,
    rootGroups,
    args.useRofi,
    ensureIconForPicker,
    normalizedSearch,
    themePath,
  );
  if (!selectedRootId) fail('No Jellyfin show/movie selected.');
  const selectedRoot = rootById.get(selectedRootId);
  if (!selectedRoot) fail('Invalid Jellyfin root selection.');

  if (isJellyfinPlayableType(selectedRoot.type)) {
    return selectedRoot.id;
  }

  const pickPlayableDescendants = async (parentId: string): Promise<string> => {
    const descendantItems = await listItemsViaApp(parentId, {
      limit: browseLimit,
      recursive: true,
      includeItemTypes: recursivePlayableIncludeItemTypes,
    });
    const playableItems = descendantItems.filter((item) => isJellyfinPlayableType(item.type));
    if (playableItems.length === 0) {
      fail('No playable Jellyfin items found.');
    }
    const selectedItemId = pickItem(
      pickerSession,
      playableItems,
      args.useRofi,
      ensureIconForPicker,
      '',
      themePath,
    );
    if (!selectedItemId) {
      fail('No Jellyfin item selected.');
    }
    return selectedItemId;
  };

  let currentContainerId = selectedRoot.id;
  while (true) {
    const directoryEntries = await listItemsViaApp(currentContainerId, {
      limit: browseLimit,
      recursive: false,
      includeItemTypes: directoryIncludeItemTypes,
    });

    const seenIds = new Set<string>();
    const childGroups: JellyfinGroupEntry[] = [];
    for (const item of directoryEntries) {
      if (!item.id || seenIds.has(item.id)) continue;
      if (!isJellyfinContainerType(item.type) && !isJellyfinPlayableType(item.type)) continue;
      seenIds.add(item.id);
      childGroups.push({
        id: item.id,
        name: item.name,
        type: item.type,
        display: `${item.name} (${item.type})`,
      });
    }

    if (childGroups.length === 0) {
      return await pickPlayableDescendants(currentContainerId);
    }

    const childById = new Map(childGroups.map((group) => [group.id, group]));
    const selectedChildId = pickGroup(
      pickerSession,
      childGroups,
      args.useRofi,
      ensureIconForPicker,
      '',
      themePath,
    );
    if (!selectedChildId) fail('No Jellyfin folder/file selected.');
    const selectedChild = childById.get(selectedChildId);
    if (!selectedChild) fail('Invalid Jellyfin item selection.');
    if (isJellyfinPlayableType(selectedChild.type)) {
      return selectedChild.id;
    }
    if (isJellyfinContainerType(selectedChild.type)) {
      return await pickPlayableDescendants(selectedChild.id);
    }
    fail('Selected Jellyfin item is not playable.');
  }
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
      id: string;
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
  const envAccessToken = (process.env.SUBMINER_JELLYFIN_ACCESS_TOKEN || '').trim();
  const envUserId = (process.env.SUBMINER_JELLYFIN_USER_ID || '').trim();
  const session: JellyfinSessionConfig = {
    serverUrl: sanitizeServerUrl(args.jellyfinServer || config.serverUrl || ''),
    accessToken: envAccessToken,
    userId: envUserId,
    defaultLibraryId: config.defaultLibraryId || '',
    pullPictures: config.pullPictures === true,
    iconCacheDir: config.iconCacheDir || '',
  };

  const rofiTheme = args.useRofi ? findRofiTheme(scriptPath) : null;
  if (args.useRofi && !rofiTheme) {
    log('warn', args.logLevel, 'Rofi theme not found for Jellyfin picker; using rofi defaults.');
  }

  const hasDirectSession = Boolean(session.serverUrl && session.accessToken && session.userId);
  let itemId = '';
  if (hasDirectSession) {
    itemId = await resolveJellyfinSelection(args, session, rofiTheme);
  } else {
    const configPath = resolveLauncherMainConfigPath();
    if (!hasStoredJellyfinSession(configPath)) {
      fail(
        'Missing Jellyfin session. Run `subminer jellyfin -l --server <url> --username <user> --password <pass>` first.',
      );
    }
    const previewAuth = await requestJellyfinPreviewAuthFromApp(appPath, args);
    if (previewAuth) {
      session.serverUrl = previewAuth.serverUrl || session.serverUrl;
      session.accessToken = previewAuth.accessToken;
      session.userId = previewAuth.userId || session.userId;
      log('debug', args.logLevel, 'Jellyfin preview auth bridge ready for picker image previews.');
    } else {
      log(
        'debug',
        args.logLevel,
        'Jellyfin preview auth bridge unavailable; picker image previews may be disabled.',
      );
    }
    itemId = await resolveJellyfinSelectionViaApp(appPath, args, session, rofiTheme);
  }
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
  const forwarded = ['--start', '--jellyfin-play', `--jellyfin-item-id=${itemId}`];
  if (args.logLevel !== 'info') forwarded.push('--log-level', args.logLevel);
  if (args.passwordStore) forwarded.push('--password-store', args.passwordStore);
  runAppCommandWithInheritLogged(appPath, forwarded, args.logLevel, 'jellyfin-play');
}

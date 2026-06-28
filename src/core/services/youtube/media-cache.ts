import { spawn as spawnProcess } from 'node:child_process';
import * as crypto from 'node:crypto';
import { EventEmitter } from 'node:events';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

import type { YoutubeMediaCacheMode } from '../../../types/integrations';
import { getYoutubeYtDlpCommand } from './ytdlp-command';

type MediaCacheSessionState = 'running' | 'ready' | 'failed';

type SpawnedProcess = EventEmitter & {
  killed?: boolean;
  kill?: (signal?: NodeJS.Signals | number) => boolean;
};

type SpawnProcess = (
  command: string,
  args: string[],
  options?: { stdio?: Array<'ignore' | 'pipe'> },
) => SpawnedProcess;

interface MediaCacheSession {
  url: string;
  dir: string;
  maxHeight: number;
  process: SpawnedProcess | null;
  readyPath: string | null;
  state: MediaCacheSessionState;
}

export interface YoutubeMediaCacheStartOptions {
  mode: YoutubeMediaCacheMode;
  maxHeight?: number;
}

export interface YoutubeMediaCacheServiceDeps {
  cacheRoot?: string;
  getYtDlpCommand?: () => string;
  spawn?: SpawnProcess;
  onDownloadStarted?: (event: { url: string }) => void;
  onReady?: (event: { url: string; path: string }) => void;
  onFailed?: (event: { url: string }) => void;
  logInfo?: (message: string) => void;
  logWarn?: (message: string) => void;
}

const MEDIA_FILE_EXTENSIONS = new Set(['.mkv', '.mp4', '.webm', '.m4a', '.mp3', '.opus']);
const DEFAULT_MAX_HEIGHT = 720;

function cacheKeyForUrl(url: string): string {
  return crypto.createHash('sha256').update(url).digest('hex').slice(0, 24);
}

function isFinalMediaFile(fileName: string): boolean {
  if (!fileName.startsWith('media.')) {
    return false;
  }
  if (fileName.endsWith('.part') || fileName.endsWith('.ytdl') || fileName.endsWith('.tmp')) {
    return false;
  }
  return MEDIA_FILE_EXTENSIONS.has(path.extname(fileName).toLowerCase());
}

function findReadyMediaPath(dir: string): string | null {
  try {
    const files = fs.readdirSync(dir);
    const mediaFile = files.find(isFinalMediaFile);
    return mediaFile ? path.join(dir, mediaFile) : null;
  } catch {
    return null;
  }
}

function getFormatSelector(maxHeight: number): string {
  return maxHeight > 0
    ? `bestvideo*[height<=${maxHeight}]+bestaudio/best[height<=${maxHeight}]`
    : 'bestvideo*+bestaudio/best';
}

function normalizeMaxHeight(maxHeight: number | undefined): number {
  if (maxHeight === undefined) {
    return DEFAULT_MAX_HEIGHT;
  }
  return Number.isInteger(maxHeight) && maxHeight >= 0 ? maxHeight : DEFAULT_MAX_HEIGHT;
}

function createYtDlpArgs(url: string, outputTemplate: string, maxHeight?: number): string[] {
  return [
    '--no-playlist',
    '--no-warnings',
    '--force-ipv4',
    '--retries',
    '5',
    '--fragment-retries',
    '5',
    '--extractor-retries',
    '5',
    '-f',
    getFormatSelector(normalizeMaxHeight(maxHeight)),
    '--merge-output-format',
    'mkv',
    '-o',
    outputTemplate,
    url,
  ];
}

export function createYoutubeMediaCacheService(deps: YoutubeMediaCacheServiceDeps = {}) {
  const cacheRoot = deps.cacheRoot ?? path.join(os.tmpdir(), 'subminer-youtube-media-cache');
  const getYtDlpCommand = deps.getYtDlpCommand ?? getYoutubeYtDlpCommand;
  const spawn: SpawnProcess =
    deps.spawn ??
    ((command, args, options) =>
      spawnProcess(command, args, options ?? {}) as unknown as SpawnedProcess);
  const sessions = new Map<string, MediaCacheSession>();
  let activeKey: string | null = null;

  const getSessionDir = (url: string): string => path.join(cacheRoot, cacheKeyForUrl(url));
  const removeCacheDir = (dir: string): void => {
    try {
      fs.rmSync(dir, { recursive: true, force: true });
    } catch {
      // Temp cache cleanup should not block shutdown or playback startup.
    }
  };
  const removeCacheRootEntriesExcept = (dirsToKeep: string[]): void => {
    let entries: fs.Dirent[];
    try {
      entries = fs.readdirSync(cacheRoot, { withFileTypes: true });
    } catch {
      return;
    }
    const keepDirs = new Set(dirsToKeep.map((dir) => path.resolve(dir)));
    for (const entry of entries) {
      const entryPath = path.join(cacheRoot, entry.name);
      if (keepDirs.has(path.resolve(entryPath))) {
        continue;
      }
      removeCacheDir(entryPath);
    }
  };
  const removeSession = (key: string): void => {
    const session = sessions.get(key);
    if (!session) {
      return;
    }
    if (session.state === 'running' && session.process?.kill && !session.process.killed) {
      session.process.kill();
    }
    sessions.delete(key);
    if (activeKey === key) {
      activeKey = null;
    }
    removeCacheDir(session.dir);
  };
  const removeInactiveSessions = (keyToKeep: string): void => {
    for (const key of [...sessions.keys()]) {
      if (key !== keyToKeep) {
        removeSession(key);
      }
    }
  };
  removeCacheRootEntriesExcept([]);

  const getCachedMediaPath = async (url: string): Promise<string | null> => {
    const key = cacheKeyForUrl(url);
    const session = sessions.get(key);
    if (session?.readyPath && fs.existsSync(session.readyPath)) {
      return session.readyPath;
    }

    const readyPath = findReadyMediaPath(session?.dir ?? getSessionDir(url));
    if (readyPath) {
      sessions.set(key, {
        url,
        dir: path.dirname(readyPath),
        maxHeight: session?.maxHeight ?? DEFAULT_MAX_HEIGHT,
        process: null,
        readyPath,
        state: 'ready',
      });
      return readyPath;
    }

    return null;
  };

  const getActiveCachedMediaPath = async (): Promise<string | null> => {
    if (!activeKey) {
      return null;
    }
    const session = sessions.get(activeKey);
    return session ? getCachedMediaPath(session.url) : null;
  };

  const start = (url: string, options: YoutubeMediaCacheStartOptions): void => {
    if (options.mode !== 'background') {
      if (activeKey) {
        const activeSession = sessions.get(activeKey);
        if (activeSession?.state === 'running') {
          removeSession(activeKey);
        }
      }
      activeKey = null;
      return;
    }

    const key = cacheKeyForUrl(url);
    const maxHeight = normalizeMaxHeight(options.maxHeight);
    activeKey = key;
    const dir = getSessionDir(url);
    const existingSession = sessions.get(key);
    const canReuseExistingSession = existingSession?.maxHeight === maxHeight;
    if (existingSession?.state === 'running' && canReuseExistingSession) {
      removeInactiveSessions(key);
      removeCacheRootEntriesExcept([existingSession.dir]);
      return;
    }
    if (existingSession) {
      if (
        canReuseExistingSession &&
        existingSession.state === 'ready' &&
        ((existingSession.readyPath && fs.existsSync(existingSession.readyPath)) ||
          findReadyMediaPath(existingSession.dir))
      ) {
        removeInactiveSessions(key);
        removeCacheRootEntriesExcept([existingSession.dir]);
        return;
      }
      removeSession(key);
      activeKey = key;
    }
    removeInactiveSessions(key);
    removeCacheRootEntriesExcept([dir]);

    fs.mkdirSync(dir, { recursive: true });
    const outputTemplate = path.join(dir, 'media.%(ext)s');
    const args = createYtDlpArgs(url, outputTemplate, maxHeight);
    const child = spawn(getYtDlpCommand(), args, { stdio: ['ignore', 'ignore', 'ignore'] });
    const session: MediaCacheSession = {
      url,
      dir,
      maxHeight,
      process: child,
      readyPath: null,
      state: 'running',
    };
    sessions.set(key, session);
    deps.logInfo?.(`Started YouTube media cache download for ${url}`);
    deps.onDownloadStarted?.({ url });

    child.once('error', (error) => {
      const currentSession = sessions.get(key);
      if (currentSession !== session) {
        return;
      }
      session.state = 'failed';
      session.process = null;
      deps.logWarn?.(
        `YouTube media cache download failed: ${
          error instanceof Error ? error.message : String(error)
        }`,
      );
      deps.onFailed?.({ url });
    });

    child.once('close', (code) => {
      const currentSession = sessions.get(key);
      if (currentSession !== session) {
        return;
      }
      if (session.state === 'failed') {
        return;
      }
      session.process = null;
      if (code === 0) {
        const readyPath = findReadyMediaPath(dir);
        if (readyPath) {
          session.state = 'ready';
          session.readyPath = readyPath;
          deps.logInfo?.(`YouTube media cache ready at ${readyPath}`);
          deps.onReady?.({ url, path: readyPath });
          return;
        }
      }
      session.state = 'failed';
      deps.logWarn?.(`YouTube media cache download exited without a usable media file.`);
      deps.onFailed?.({ url });
    });
  };

  const cleanup = (): void => {
    for (const key of [...sessions.keys()]) {
      removeSession(key);
    }
    activeKey = null;
  };

  return {
    cleanup,
    getActiveCachedMediaPath,
    getCachedMediaPath,
    start,
  };
}

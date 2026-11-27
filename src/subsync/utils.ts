import * as fs from 'fs';
import * as childProcess from 'child_process';
import * as path from 'path';
import { DEFAULT_CONFIG } from '../config';
import {
  buildWindowsShimInvocation,
  isWindowsCommandShim,
} from '../core/utils/windows-command-shim';
import { SubsyncConfig, SubsyncMode } from '../types';

export interface MpvTrack {
  id?: number;
  type?: string;
  selected?: boolean;
  external?: boolean;
  forced?: boolean;
  default?: boolean;
  lang?: string;
  title?: string;
  codec?: string;
  'ff-index'?: number;
  'external-filename'?: string;
}

export interface SubsyncResolvedConfig {
  mode?: SubsyncMode;
  alassPath: string;
  ffsubsyncPath: string;
  ffmpegPath: string;
  replace?: boolean;
}

const DEFAULT_SUBSYNC_EXECUTABLE_PATHS = {
  alass: '/usr/bin/alass',
  ffsubsync: '/usr/bin/ffsubsync',
  ffmpeg: '/usr/bin/ffmpeg',
} as const;

const COMMON_BIN_DIRS = ['/opt/homebrew/bin', '/usr/local/bin', '/usr/bin'];
const WINDOWS_RUNNABLE_EXTENSIONS = new Set(['.exe', '.com', '.cmd', '.bat', '.ps1', '.sh']);

function isExecutablePath(filePath: string): boolean {
  try {
    if (!fs.statSync(filePath).isFile()) return false;
    if (process.platform !== 'win32') {
      fs.accessSync(filePath, fs.constants.X_OK);
    }
    return true;
  } catch {
    return false;
  }
}

export function candidateExecutableNames(
  name: string,
  platform: NodeJS.Platform = process.platform,
  pathExt: string | undefined = process.env.PATHEXT,
): string[] {
  if (platform !== 'win32' || path.extname(name)) return [name];
  const extensions = (pathExt ?? '.COM;.EXE;.BAT;.CMD')
    .split(';')
    .map((ext) => ext.trim())
    .filter((ext) => ext.length > 0 && WINDOWS_RUNNABLE_EXTENSIONS.has(ext.toLowerCase()));
  const effective = extensions.length > 0 ? extensions : ['.exe', '.com', '.cmd', '.bat'];
  return effective.map((ext) => `${name}${ext}`);
}

/**
 * Find an executable on PATH (or common install dirs), honoring an explicit
 * configured override first. Returns null if not found.
 */
export function findSubsyncExecutable(
  configuredPath: string | undefined,
  names: readonly string[],
): string | null {
  const trimmed = configuredPath?.trim() ?? '';
  if (trimmed) {
    if (hasPathSeparators(trimmed)) {
      return isExecutablePath(trimmed) ? trimmed : null;
    }
    const foundConfigured = findSubsyncExecutable(undefined, [trimmed]);
    return foundConfigured;
  }

  const dirs = [
    ...(process.env.PATH ?? '')
      .split(path.delimiter)
      .map((entry) => entry.trim())
      .filter(Boolean),
    ...(process.platform === 'win32' ? [] : COMMON_BIN_DIRS),
  ];

  const seen = new Set<string>();
  for (const dir of dirs) {
    if (seen.has(dir)) continue;
    seen.add(dir);
    for (const name of names) {
      for (const candidateName of candidateExecutableNames(name)) {
        const candidate = path.join(dir, candidateName);
        if (isExecutablePath(candidate)) return candidate;
      }
    }
  }
  return null;
}

export interface SubsyncContext {
  videoPath: string;
  workingDirectory?: string | null;
  durationSec?: number | null;
  primaryTrack: MpvTrack;
  secondaryTrack: MpvTrack | null;
  /** Every usable subtitle track, including the primary one. */
  subtitleTracks: MpvTrack[];
  audioStreamIndex: number | null;
}

export interface CommandResult {
  ok: boolean;
  code: number | null;
  stderr: string;
  stdout: string;
  error?: string;
}

export function resolveCommandInvocation(
  executable: string,
  args: string[],
  platform: NodeJS.Platform = process.platform,
): { command: string; args: string[]; windowsVerbatimArguments?: boolean } {
  if (platform !== 'win32') {
    return { command: executable, args };
  }

  const normalizeBashArg = (value: string): string => {
    const normalized = value.replace(/\\/g, '/');
    const driveMatch = normalized.match(/^([A-Za-z]):\/(.*)$/);
    if (!driveMatch) {
      return normalized;
    }

    const [, driveLetter, remainder] = driveMatch;
    return `/mnt/${driveLetter!.toLowerCase()}/${remainder}`;
  };
  const extension = path.extname(executable).toLowerCase();
  if (isWindowsCommandShim(executable)) {
    const shim = buildWindowsShimInvocation(executable, args, process.env);
    return {
      command: shim.file,
      args: shim.args,
      windowsVerbatimArguments: true,
    };
  }
  if (extension === '.ps1') {
    return {
      command: 'powershell.exe',
      args: ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', executable, ...args],
    };
  }

  if (extension === '.sh') {
    return {
      command: 'bash',
      args: [normalizeBashArg(executable), ...args.map(normalizeBashArg)],
    };
  }

  return { command: executable, args };
}

export function getSubsyncConfig(config: SubsyncConfig | undefined): SubsyncResolvedConfig {
  const resolvePath = (
    value: string | undefined,
    names: readonly string[],
    fallback: string,
  ): string => {
    const trimmed = value?.trim();
    if (trimmed && trimmed.length > 0) {
      if (hasPathSeparators(trimmed)) return trimmed;
      return findSubsyncExecutable(trimmed, [trimmed]) ?? trimmed;
    }
    return findSubsyncExecutable(undefined, names) ?? fallback;
  };

  return {
    mode: config?.mode ?? DEFAULT_CONFIG.subsync.mode,
    alassPath: resolvePath(
      config?.alass_path,
      ['alass', 'alass-cli'],
      DEFAULT_SUBSYNC_EXECUTABLE_PATHS.alass,
    ),
    ffsubsyncPath: resolvePath(
      config?.ffsubsync_path,
      ['ffsubsync'],
      DEFAULT_SUBSYNC_EXECUTABLE_PATHS.ffsubsync,
    ),
    ffmpegPath: resolvePath(
      config?.ffmpeg_path,
      ['ffmpeg'],
      DEFAULT_SUBSYNC_EXECUTABLE_PATHS.ffmpeg,
    ),
    replace: config?.replace ?? DEFAULT_CONFIG.subsync.replace,
  };
}

export function hasPathSeparators(value: string): boolean {
  return value.includes('/') || value.includes('\\');
}

export function fileExists(pathOrEmpty: string): boolean {
  if (!pathOrEmpty) return false;
  try {
    return fs.existsSync(pathOrEmpty);
  } catch {
    return false;
  }
}

export function formatTrackLabel(track: MpvTrack): string {
  const trackId = typeof track.id === 'number' ? track.id : -1;
  const source = track.external ? 'External' : 'Internal';
  const lang = track.lang || track.title || 'unknown';
  const active = track.selected ? ' (active)' : '';
  return `${source} #${trackId} - ${lang}${active}`;
}

export function getTrackById(tracks: MpvTrack[], trackId: number | null): MpvTrack | null {
  if (trackId === null) return null;
  return tracks.find((track) => track.id === trackId) ?? null;
}

export function codecToExtension(codec: string | undefined): string | null {
  if (!codec) return null;
  const normalized = codec.toLowerCase();
  if (
    normalized === 'subrip' ||
    normalized === 'srt' ||
    normalized === 'text' ||
    normalized === 'mov_text'
  )
    return 'srt';
  if (normalized === 'ass' || normalized === 'ssa') return 'ass';
  if (normalized === 'webvtt' || normalized === 'vtt') return 'vtt';
  if (normalized === 'ttml') return 'ttml';
  return null;
}

export function extensionToFfmpegMuxer(extension: string): string {
  if (extension === 'vtt') return 'webvtt';
  return extension;
}

/**
 * Derive Jellyfin's direct subtitle delivery endpoint for an embedded track on a
 * direct-play stream URL (`.../Videos/{itemId}/stream?static=true&MediaSourceId=...&ApiKey=...`).
 * Returns null for non-Jellyfin or transcoded URLs where stream indices may not match.
 */
export function buildJellyfinEmbeddedSubtitleUrl(
  videoPath: string,
  ffIndex: number,
  extension: string,
): string | null {
  if (!Number.isInteger(ffIndex) || ffIndex < 0 || !extension) return null;
  let url: URL;
  try {
    url = new URL(videoPath);
  } catch {
    return null;
  }
  if (url.protocol !== 'http:' && url.protocol !== 'https:') return null;

  const findParam = (pattern: RegExp): string | null => {
    for (const [key, value] of url.searchParams) {
      if (pattern.test(key) && value.trim().length > 0) {
        return value.trim();
      }
    }
    return null;
  };

  if (findParam(/^static$/i)?.toLowerCase() !== 'true') return null;

  const apiKey = findParam(/^api_?key$/i);
  const mediaSourceId = findParam(/^mediasourceid$/i);
  if (!apiKey || !mediaSourceId) return null;

  const match = /^(.*)\/videos\/([^/]+)\/stream(?:\.[^/]+)?$/i.exec(url.pathname);
  if (!match) return null;

  const basePath = match[1] ?? '';
  const rawItemId = match[2] ?? '';
  if (!rawItemId) return null;

  let itemId = rawItemId;
  try {
    itemId = decodeURIComponent(rawItemId);
  } catch {
    // Keep raw segment if it contains malformed percent escapes.
  }

  const deliveryExt = extension === 'ttml' ? 'srt' : extension;
  const subtitleUrl = new URL(url.origin);
  subtitleUrl.pathname = `${basePath}/Videos/${encodeURIComponent(itemId)}/${encodeURIComponent(mediaSourceId)}/Subtitles/${ffIndex}/Stream.${deliveryExt}`;
  subtitleUrl.searchParams.set('ApiKey', apiKey);
  return subtitleUrl.toString();
}

export function runCommand(
  executable: string,
  args: string[],
  timeoutMs = 120000,
): Promise<CommandResult> {
  return new Promise((resolve) => {
    const invocation = resolveCommandInvocation(executable, args);
    const child = childProcess.spawn(invocation.command, invocation.args, {
      stdio: ['ignore', 'pipe', 'pipe'],
      ...(invocation.windowsVerbatimArguments ? { windowsVerbatimArguments: true } : {}),
    });
    let stdout = '';
    let stderr = '';
    const timeout = setTimeout(() => {
      child.kill('SIGKILL');
    }, timeoutMs);

    child.stdout.on('data', (chunk: Buffer) => {
      stdout += chunk.toString();
    });
    child.stderr.on('data', (chunk: Buffer) => {
      stderr += chunk.toString();
    });
    child.on('error', (error: Error) => {
      clearTimeout(timeout);
      resolve({
        ok: false,
        code: null,
        stderr,
        stdout,
        error: error.message,
      });
    });
    child.on('close', (code: number | null) => {
      clearTimeout(timeout);
      resolve({
        ok: code === 0,
        code,
        stderr,
        stdout,
      });
    });
  });
}

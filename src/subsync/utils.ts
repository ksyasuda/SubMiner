import * as fs from 'fs';
import * as childProcess from 'child_process';
import * as path from 'path';
import { DEFAULT_CONFIG } from '../config';
import { SubsyncConfig, SubsyncMode } from '../types';

export interface MpvTrack {
  id?: number;
  type?: string;
  selected?: boolean;
  external?: boolean;
  lang?: string;
  title?: string;
  codec?: string;
  'ff-index'?: number;
  'external-filename'?: string;
}

export interface SubsyncResolvedConfig {
  defaultMode: SubsyncMode;
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

export interface SubsyncContext {
  videoPath: string;
  primaryTrack: MpvTrack;
  secondaryTrack: MpvTrack | null;
  sourceTracks: MpvTrack[];
  audioStreamIndex: number | null;
}

export interface CommandResult {
  ok: boolean;
  code: number | null;
  stderr: string;
  stdout: string;
  error?: string;
}

function resolveCommandInvocation(
  executable: string,
  args: string[],
): { command: string; args: string[] } {
  if (process.platform !== 'win32') {
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
  const resolvePath = (value: string | undefined, fallback: string): string => {
    const trimmed = value?.trim();
    return trimmed && trimmed.length > 0 ? trimmed : fallback;
  };

  return {
    defaultMode: config?.defaultMode ?? DEFAULT_CONFIG.subsync.defaultMode,
    alassPath: resolvePath(config?.alass_path, DEFAULT_SUBSYNC_EXECUTABLE_PATHS.alass),
    ffsubsyncPath: resolvePath(config?.ffsubsync_path, DEFAULT_SUBSYNC_EXECUTABLE_PATHS.ffsubsync),
    ffmpegPath: resolvePath(config?.ffmpeg_path, DEFAULT_SUBSYNC_EXECUTABLE_PATHS.ffmpeg),
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

export function runCommand(
  executable: string,
  args: string[],
  timeoutMs = 120000,
): Promise<CommandResult> {
  return new Promise((resolve) => {
    const invocation = resolveCommandInvocation(executable, args);
    const child = childProcess.spawn(invocation.command, invocation.args, {
      stdio: ['ignore', 'pipe', 'pipe'],
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

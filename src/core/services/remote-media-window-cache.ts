import { execFile as nodeExecFile, type ExecFileException } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { createLogger } from '../../logger';
import { normalizeMediaInput, type MediaInput, type MediaInputOptions } from '../../media-input';

const log = createLogger('media-window');

export const REMOTE_MEDIA_WINDOW_TIMEOUT_MS = 120_000;
export const REMOTE_MEDIA_WINDOW_MAX_SECONDS = 180;
const HEAD_SLACK_SECONDS = 0.25;
const TAIL_SLACK_SECONDS = 1;
const DEFAULT_IDLE_TTL_MS = 10 * 60_000;
const COVERAGE_EPSILON_SECONDS = 0.01;

export interface RemoteMediaWindowSource {
  path: string;
  inputOptions?: MediaInputOptions;
  /** FFmpeg stream index to keep; `null`/undefined keeps every audio stream. */
  audioStreamIndex?: number | null;
}

export interface RemoteMediaWindowRange {
  startTime: number;
  endTime: number;
}

export interface RemoteMediaWindow {
  path: string;
  startTime: number;
  endTime: number;
  sourcePath: string;
  audioStreamIndex: number | null;
  /** Input descriptor for FFmpeg reads; timestamps stay absolute so callers keep source times. */
  media: MediaInput;
}

type WindowExecFile = (
  file: string,
  args: readonly string[],
  options: { timeout: number },
  callback: (error: ExecFileException | null) => void,
) => void;

export interface RemoteMediaWindowCacheOptions {
  tempDir?: string;
  execFile?: WindowExecFile;
  idleTtlMs?: number;
  logDebug?: (message: string) => void;
}

interface PendingFetch extends RemoteMediaWindowRange {
  sourcePath: string;
  audioStreamIndex: number | null;
  promise: Promise<RemoteMediaWindow>;
}

export function isRemoteMediaWindowSourcePath(value: string): boolean {
  return /^https?:\/\//i.test(value.trim());
}

function describeSourceForDebugLog(sourcePath: string): string {
  try {
    return `remote:${new URL(sourcePath).hostname.toLowerCase() || 'unknown'}`;
  } catch {
    return 'remote:unknown';
  }
}

function isUsableRange(range: RemoteMediaWindowRange, allowEmpty: boolean): boolean {
  return (
    Number.isFinite(range.startTime) &&
    Number.isFinite(range.endTime) &&
    range.startTime >= 0 &&
    (allowEmpty ? range.endTime >= range.startTime : range.endTime > range.startTime)
  );
}

function audioStreamMatches(
  windowIndex: number | null,
  requested: number | null | undefined,
): boolean {
  return requested == null || windowIndex === requested;
}

function covers(
  candidate: RemoteMediaWindowRange & { sourcePath: string; audioStreamIndex: number | null },
  source: RemoteMediaWindowSource,
  range: RemoteMediaWindowRange,
): boolean {
  return (
    candidate.sourcePath === source.path &&
    audioStreamMatches(candidate.audioStreamIndex, source.audioStreamIndex) &&
    candidate.startTime <= range.startTime + COVERAGE_EPSILON_SECONDS &&
    candidate.endTime >= range.endTime - COVERAGE_EPSILON_SECONDS
  );
}

/**
 * Stream-copies `[startTime, endTime]` of a remote source into a local Matroska file.
 * `-copyts -start_at_zero` keeps the source timestamps, so later reads seek with the
 * original times via `-seek_timestamp 1` (see `MediaInput.absoluteTimestamps`).
 */
export function buildRemoteMediaWindowArgs(
  source: RemoteMediaWindowSource,
  range: RemoteMediaWindowRange,
  outputPath: string,
): string[] {
  const input = normalizeMediaInput({ path: source.path, inputOptions: source.inputOptions });
  const audioMap =
    typeof source.audioStreamIndex === 'number' && Number.isInteger(source.audioStreamIndex)
      ? `0:${source.audioStreamIndex}`
      : '0:a';
  return [
    '-hide_banner',
    '-nostdin',
    '-loglevel',
    'error',
    '-ss',
    String(range.startTime),
    '-t',
    String(range.endTime - range.startTime),
    ...input.inputArgs,
    '-i',
    input.path,
    '-map',
    '0:v:0?',
    '-map',
    audioMap,
    '-c',
    'copy',
    '-sn',
    '-dn',
    '-copyts',
    '-start_at_zero',
    '-f',
    'matroska',
    '-y',
    outputPath,
  ];
}

/**
 * Holds one downloaded window of the current remote stream so the timing review,
 * audio extraction, and screenshot all read the same local bytes instead of each
 * re-fetching the clip over HTTP. A new window replaces the old one; the file is
 * deleted after `idleTtlMs` without use, on `clear()`, or on `cleanup()`.
 */
export class RemoteMediaWindowCache {
  private readonly tempDir: string;
  private readonly execFile: WindowExecFile;
  private readonly idleTtlMs: number;
  private readonly logDebug: (message: string) => void;
  private current: RemoteMediaWindow | null = null;
  private pending: PendingFetch | null = null;
  private idleTimer: ReturnType<typeof setTimeout> | null = null;
  private sequence = 0;

  constructor(options: RemoteMediaWindowCacheOptions = {}) {
    this.tempDir = options.tempDir ?? path.join(os.tmpdir(), 'subminer-media-windows');
    this.execFile = options.execFile ?? nodeExecFile;
    this.idleTtlMs = options.idleTtlMs ?? DEFAULT_IDLE_TTL_MS;
    this.logDebug = options.logDebug ?? ((message) => log.debug(message));
  }

  get currentWindow(): RemoteMediaWindow | null {
    return this.current;
  }

  /** Returns a ready or in-flight window covering the range; never starts a download. */
  async lookup(
    source: RemoteMediaWindowSource,
    range: RemoteMediaWindowRange,
  ): Promise<RemoteMediaWindow | null> {
    if (!isUsableRange(range, true)) return null;
    if (this.current && covers(this.current, source, range)) {
      this.touch();
      return this.current;
    }
    const pending = this.pending;
    if (pending && covers(pending, source, range)) {
      try {
        const window = await pending.promise;
        this.touch();
        return window;
      } catch {
        return null;
      }
    }
    return null;
  }

  /** Returns a window covering the range, downloading (and widening) one when needed. */
  async acquire(
    source: RemoteMediaWindowSource,
    range: RemoteMediaWindowRange,
  ): Promise<RemoteMediaWindow> {
    if (!isUsableRange(range, false)) {
      throw new Error('Media window range is invalid.');
    }
    if (range.endTime - range.startTime > REMOTE_MEDIA_WINDOW_MAX_SECONDS) {
      throw new Error('Media window range is too long to download.');
    }

    for (;;) {
      const hit = await this.lookup(source, range);
      if (hit) return hit;
      const pending = this.pending;
      if (!pending) break;
      // Another caller is already downloading; wait for it, then re-check coverage.
      await pending.promise.catch(() => null);
    }

    return this.fetch(source, this.planFetchRange(source, range));
  }

  clear(): void {
    this.cancelIdleTimer();
    const current = this.current;
    this.current = null;
    if (current) this.removeFile(current.path);
  }

  cleanup(): void {
    this.clear();
    try {
      fs.rmSync(this.tempDir, { recursive: true, force: true });
    } catch (error) {
      log.error('Failed to cleanup media window directory:', error);
    }
  }

  private planFetchRange(
    source: RemoteMediaWindowSource,
    range: RemoteMediaWindowRange,
  ): RemoteMediaWindowRange {
    let startTime = Math.max(0, range.startTime - HEAD_SLACK_SECONDS);
    let endTime = range.endTime + TAIL_SLACK_SECONDS;
    const current = this.current;
    if (
      current &&
      current.sourcePath === source.path &&
      audioStreamMatches(current.audioStreamIndex, source.audioStreamIndex)
    ) {
      // Keep what was already downloaded when the review timeline grows in one direction.
      const unionStart = Math.min(startTime, current.startTime);
      const unionEnd = Math.max(endTime, current.endTime);
      if (unionEnd - unionStart <= REMOTE_MEDIA_WINDOW_MAX_SECONDS) {
        startTime = unionStart;
        endTime = unionEnd;
      }
    }
    return { startTime, endTime };
  }

  private fetch(
    source: RemoteMediaWindowSource,
    range: RemoteMediaWindowRange,
  ): Promise<RemoteMediaWindow> {
    fs.mkdirSync(this.tempDir, { recursive: true });
    this.sequence += 1;
    const outputPath = path.join(this.tempDir, `window_${Date.now()}_${this.sequence}.mkv`);
    const audioStreamIndex =
      typeof source.audioStreamIndex === 'number' ? source.audioStreamIndex : null;
    const description = describeSourceForDebugLog(source.path);
    const startedAt = Date.now();
    this.logDebug(
      `[media-window] fetch start ${description} start=${range.startTime} end=${range.endTime} audioStream=${audioStreamIndex ?? 'all'}`,
    );

    const promise = new Promise<RemoteMediaWindow>((resolve, reject) => {
      this.execFile(
        'ffmpeg',
        buildRemoteMediaWindowArgs(source, range, outputPath),
        { timeout: REMOTE_MEDIA_WINDOW_TIMEOUT_MS },
        (error) => {
          const elapsedMs = Math.max(0, Date.now() - startedAt);
          const size = error ? 0 : this.fileSize(outputPath);
          if (error || size === 0) {
            this.removeFile(outputPath);
            const reason = error
              ? error.code === 'ENOENT'
                ? 'FFmpeg not found. Install FFmpeg to enable media generation.'
                : `FFmpeg media window failed: ${error.message}`
              : 'FFmpeg exited without creating a media window.';
            this.logDebug(`[media-window] fetch failed ${description} elapsedMs=${elapsedMs}`);
            reject(new Error(reason));
            return;
          }
          const window: RemoteMediaWindow = {
            path: outputPath,
            startTime: range.startTime,
            endTime: range.endTime,
            sourcePath: source.path,
            audioStreamIndex,
            media: {
              path: outputPath,
              source: 'remote-window',
              singleResolvedStream: true,
              absoluteTimestamps: true,
            },
          };
          this.logDebug(
            `[media-window] fetch complete ${description} elapsedMs=${elapsedMs} bytes=${size}`,
          );
          this.replaceCurrent(window);
          resolve(window);
        },
      );
    });

    const pending: PendingFetch = {
      sourcePath: source.path,
      audioStreamIndex,
      startTime: range.startTime,
      endTime: range.endTime,
      promise,
    };
    this.pending = pending;
    promise
      .catch(() => undefined)
      .then(() => {
        if (this.pending === pending) this.pending = null;
      });
    return promise;
  }

  private replaceCurrent(window: RemoteMediaWindow): void {
    const previous = this.current;
    this.current = window;
    if (previous && previous.path !== window.path) this.removeFile(previous.path);
    this.touch();
  }

  private touch(): void {
    this.cancelIdleTimer();
    if (this.idleTtlMs <= 0 || !this.current) return;
    const timer = setTimeout(() => {
      if (this.idleTimer === timer) this.idleTimer = null;
      this.clear();
    }, this.idleTtlMs);
    timer.unref?.();
    this.idleTimer = timer;
  }

  private cancelIdleTimer(): void {
    if (this.idleTimer) clearTimeout(this.idleTimer);
    this.idleTimer = null;
  }

  private fileSize(filePath: string): number {
    try {
      return fs.statSync(filePath).size;
    } catch {
      return 0;
    }
  }

  private removeFile(filePath: string): void {
    try {
      fs.unlinkSync(filePath);
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== 'ENOENT') {
        log.debug(`Failed to remove media window ${filePath}:`, (error as Error).message);
      }
    }
  }
}

let sharedCache: RemoteMediaWindowCache | null = null;

/** Process-wide cache so the review modal and card media generation share one download. */
export function getSharedRemoteMediaWindowCache(): RemoteMediaWindowCache {
  sharedCache ??= new RemoteMediaWindowCache();
  return sharedCache;
}

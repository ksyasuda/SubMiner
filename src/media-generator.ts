/*
 * SubMiner - Subtitle mining overlay for mpv
 * Copyright (C) 2024 sudacode
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

import { ExecFileException, execFile } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { createLogger } from './logger';
import { normalizeMediaInput, type MediaInput } from './media-input';

const log = createLogger('media');
const AUDIO_NORMALIZATION_FILTER = 'loudnorm=I=-23:TP=-2:LRA=11';
const AUDIO_AMPLIFICATION_LIMITER_FILTER = 'alimiter=limit=0.891251:level=false';
export const AUDIO_GENERATION_TIMEOUT_MS = 120_000;

export type { MediaInput, MediaInputOptions } from './media-input';

function normalizeAnimatedImageFps(fps: number | undefined): number {
  const fallbackFps = 10;
  const safeFps = typeof fps === 'number' && Number.isFinite(fps) ? fps : fallbackFps;
  return Math.max(1, Math.min(60, safeFps));
}

function roundDurationUpToNextFrameBoundary(seconds: number, fps: number): number {
  if (!(Number.isFinite(seconds) && seconds > 0)) {
    return 0;
  }

  return (Math.floor(seconds * fps + 1e-9) + 1) / fps;
}

export function buildAnimatedImageVideoFilter(options: {
  fps?: number;
  maxWidth?: number;
  maxHeight?: number;
  leadingStillDuration?: number;
}): string {
  const { fps = 10, maxWidth = 640, maxHeight, leadingStillDuration = 0 } = options;
  const clampedFps = normalizeAnimatedImageFps(fps);
  const alignedLeadingStillDuration = roundDurationUpToNextFrameBoundary(
    leadingStillDuration,
    clampedFps,
  );
  const vfParts: string[] = [];

  if (alignedLeadingStillDuration > 0) {
    vfParts.push(`tpad=start_duration=${alignedLeadingStillDuration}:start_mode=clone`);
  }

  vfParts.push(`fps=${clampedFps}`);

  if (maxWidth && maxWidth > 0 && maxHeight && maxHeight > 0) {
    vfParts.push(`scale=w=${maxWidth}:h=${maxHeight}:force_original_aspect_ratio=decrease`);
  } else if (maxWidth && maxWidth > 0) {
    vfParts.push(`scale=w=${maxWidth}:h=-2`);
  } else if (maxHeight && maxHeight > 0) {
    vfParts.push(`scale=w=-2:h=${maxHeight}`);
  }

  return vfParts.join(',');
}

export interface MediaGeneratorOptions {
  logDebug?: (message: string) => void;
  now?: () => number;
}

function sanitizeDebugToken(value: string, fallback: string): string {
  const trimmed = value.trim();
  if (!trimmed) {
    return fallback;
  }
  const sanitized = trimmed.replace(/[^A-Za-z0-9_.:-]+/g, '-').slice(0, 80);
  return sanitized || fallback;
}

function describeMediaInputPathForDebugLog(value: string): string {
  try {
    const url = new URL(value);
    if (url.protocol === 'http:' || url.protocol === 'https:') {
      return `remote:${url.hostname.toLowerCase() || 'unknown'}`;
    }
    return `${url.protocol.replace(/:$/, '')}:`;
  } catch {
    // Not a URL; treat as a local file path below.
  }

  if (value.startsWith('edl://')) {
    return 'edl:';
  }

  return `local:${value}`;
}

function describeMediaInputForDebugLog(input: MediaInput): string {
  const pathValue = typeof input === 'string' ? input : input.path;
  const sourceValue = typeof input === 'string' ? 'raw' : input.source;
  const source = sanitizeDebugToken(sourceValue ?? 'raw', 'raw');
  return `source=${source} input=${describeMediaInputPathForDebugLog(pathValue)}`;
}

function describeFfmpegFailureForDebugLog(error: ExecFileException): string {
  const code = typeof error.code === 'string' || typeof error.code === 'number' ? error.code : null;
  const signal = typeof error.signal === 'string' ? error.signal : null;
  if (code !== null) {
    return `code=${code}`;
  }
  if (signal) {
    return `signal=${signal}`;
  }
  return `name=${sanitizeDebugToken(error.name || 'Error', 'Error')}`;
}

export class MediaGenerator {
  private tempDir: string;
  private notifyIconDir: string;
  private av1EncoderPromise: Promise<string | null> | null = null;
  private readonly options: MediaGeneratorOptions;

  constructor(tempDir?: string, options: MediaGeneratorOptions = {}) {
    this.options = options;
    this.tempDir = tempDir || path.join(os.tmpdir(), 'subminer-media');
    this.notifyIconDir = path.join(os.tmpdir(), 'subminer-notify');
    this.ensureDirectory(this.tempDir);
    this.ensureDirectory(this.notifyIconDir);
    // Clean up old notification icons on startup (older than 1 hour)
    this.cleanupOldNotificationIcons();
  }

  private nowMs(): number {
    try {
      const value = this.options.now?.() ?? Date.now();
      return Number.isFinite(value) ? value : Date.now();
    } catch {
      return Date.now();
    }
  }

  private elapsedMs(startedAt: number): number {
    return Math.max(0, Math.round(this.nowMs() - startedAt));
  }

  private logMediaDebug(message: string): void {
    const logDebug = this.options.logDebug ?? ((line: string) => log.debug(line));
    try {
      logDebug(`[media-generator] ${message}`);
    } catch {
      // Debug logging should not affect media generation.
    }
  }

  private ensureDirectory(dir: string): void {
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
  }

  private createTempOutputPath(prefix: string, extension: string): string {
    this.ensureDirectory(this.tempDir);
    return path.join(this.tempDir, `${prefix}_${Date.now()}.${extension}`);
  }

  /**
   * Clean up notification icons older than 1 hour.
   * Called on startup to prevent accumulation of temp files.
   */
  private cleanupOldNotificationIcons(): void {
    try {
      if (!fs.existsSync(this.notifyIconDir)) return;

      const files = fs.readdirSync(this.notifyIconDir);
      const oneHourAgo = Date.now() - 60 * 60 * 1000;

      for (const file of files) {
        if (!file.endsWith('.png')) continue;
        const filePath = path.join(this.notifyIconDir, file);
        try {
          const stat = fs.statSync(filePath);
          if (stat.mtimeMs < oneHourAgo) {
            fs.unlinkSync(filePath);
          }
        } catch (err) {
          log.debug(`Failed to clean up ${filePath}:`, (err as Error).message);
        }
      }
    } catch (err) {
      log.error('Failed to cleanup old notification icons:', err);
    }
  }

  /**
   * Write a notification icon buffer to a temp file and return the file path.
   * The file path can be passed directly to Electron Notification for better
   * compatibility with Linux/Wayland notification daemons.
   */
  writeNotificationIconToFile(iconBuffer: Buffer, noteId: number): string {
    this.ensureDirectory(this.notifyIconDir);
    const filename = `icon_${noteId}_${Date.now()}.png`;
    const filePath = path.join(this.notifyIconDir, filename);
    fs.writeFileSync(filePath, iconBuffer);
    return filePath;
  }

  scheduleNotificationIconCleanup(filePath: string, delayMs = 10000): void {
    setTimeout(() => {
      try {
        fs.unlinkSync(filePath);
      } catch {}
    }, delayMs);
  }

  private ffmpegError(label: string, error: ExecFileException): Error {
    if (error.code === 'ENOENT') {
      return new Error('FFmpeg not found. Install FFmpeg to enable media generation.');
    }
    return new Error(`FFmpeg ${label} failed: ${error.message}`);
  }

  private detectAv1Encoder(): Promise<string | null> {
    if (this.av1EncoderPromise) return this.av1EncoderPromise;

    this.av1EncoderPromise = new Promise((resolve) => {
      execFile(
        'ffmpeg',
        ['-hide_banner', '-encoders'],
        { timeout: 10000 },
        (error, stdout, stderr) => {
          if (error) {
            resolve(null);
            return;
          }

          const output = `${stdout || ''}\n${stderr || ''}`;
          const candidates = ['libaom-av1', 'libsvtav1', 'librav1e'];
          for (const encoder of candidates) {
            if (output.includes(encoder)) {
              resolve(encoder);
              return;
            }
          }
          resolve(null);
        },
      );
    });

    return this.av1EncoderPromise;
  }

  async generateAudio(
    videoPath: MediaInput,
    startTime: number,
    endTime: number,
    padding: number = 0,
    audioStreamIndex: number | null = null,
    normalizeAudio = true,
    volumeScale?: number,
  ): Promise<Buffer> {
    const safePadding = Number.isFinite(padding) ? Math.max(0, padding) : 0;
    const start = Math.max(0, startTime - safePadding);
    const duration = endTime - start + safePadding;
    const mediaInput = normalizeMediaInput(videoPath);
    const inputDescription = describeMediaInputForDebugLog(videoPath);
    const hasSelectedAudioStream =
      !mediaInput.singleResolvedStream &&
      typeof audioStreamIndex === 'number' &&
      Number.isInteger(audioStreamIndex) &&
      audioStreamIndex >= 0;
    const isLocalMatroskaMedia =
      !/^[A-Za-z][A-Za-z\d+.-]*:\/\//.test(mediaInput.path) &&
      /\.(?:mkv|mka|mks|webm)$/i.test(mediaInput.path);

    return new Promise((resolve, reject) => {
      const outputPath = this.createTempOutputPath('audio', 'mp3');
      const startedAt = this.nowMs();
      const args: string[] = [
        '-ss',
        start.toString(),
        '-t',
        duration.toString(),
        ...mediaInput.inputArgs,
        ...(hasSelectedAudioStream && isLocalMatroskaMedia
          ? ['-probesize', '32768', '-analyzeduration', '0']
          : []),
        '-i',
        mediaInput.path,
      ];

      if (hasSelectedAudioStream) {
        args.push('-map', `0:${audioStreamIndex}`);
      }

      args.push('-vn');
      const audioFilters: string[] = [];
      if (normalizeAudio) {
        audioFilters.push(AUDIO_NORMALIZATION_FILTER);
      }
      if (
        typeof volumeScale === 'number' &&
        Number.isFinite(volumeScale) &&
        volumeScale >= 0 &&
        volumeScale !== 1
      ) {
        audioFilters.push(`volume=${volumeScale}`);
        if (volumeScale > 1) {
          audioFilters.push(AUDIO_AMPLIFICATION_LIMITER_FILTER);
        }
      }
      if (audioFilters.length > 0) {
        args.push('-af', audioFilters.join(','));
      }
      args.push('-acodec', 'libmp3lame', '-q:a', '2', '-ar', '44100', '-y', outputPath);

      this.logMediaDebug(
        `audio start ${inputDescription} start=${start} duration=${duration} padding=${safePadding}`,
      );
      execFile('ffmpeg', args, { timeout: AUDIO_GENERATION_TIMEOUT_MS }, (error) => {
        if (error) {
          this.logMediaDebug(
            `audio failed ${inputDescription} elapsedMs=${this.elapsedMs(startedAt)} ${describeFfmpegFailureForDebugLog(error)}`,
          );
          reject(this.ffmpegError('audio generation', error));
          return;
        }

        try {
          const data = fs.readFileSync(outputPath);
          fs.unlinkSync(outputPath);
          this.logMediaDebug(
            `audio complete ${inputDescription} elapsedMs=${this.elapsedMs(startedAt)} bytes=${data.byteLength}`,
          );
          resolve(data);
        } catch (err) {
          if ((err as NodeJS.ErrnoException).code === 'ENOENT') {
            reject(
              new Error(
                'FFmpeg audio generation failed: FFmpeg exited without creating an output file.',
              ),
            );
          } else {
            reject(err);
          }
        }
      });
    });
  }

  async generateScreenshot(
    videoPath: MediaInput,
    timestamp: number,
    options: {
      format: 'jpg' | 'png' | 'webp';
      quality?: number;
      maxWidth?: number;
      maxHeight?: number;
    },
  ): Promise<Buffer> {
    const { format, quality = 92, maxWidth, maxHeight } = options;
    const ext = format === 'webp' ? 'webp' : format === 'png' ? 'png' : 'jpg';
    const codecMap: Record<'jpg' | 'png' | 'webp', string> = {
      jpg: 'mjpeg',
      png: 'png',
      webp: 'webp',
    };
    const mediaInput = normalizeMediaInput(videoPath);
    const inputDescription = describeMediaInputForDebugLog(videoPath);

    const args: string[] = [
      '-ss',
      timestamp.toString(),
      ...mediaInput.inputArgs,
      '-i',
      mediaInput.path,
      '-vframes',
      '1',
    ];

    const vfParts: string[] = [];
    if (maxWidth && maxWidth > 0 && maxHeight && maxHeight > 0) {
      vfParts.push(`scale=w=${maxWidth}:h=${maxHeight}:force_original_aspect_ratio=decrease`);
    } else if (maxWidth && maxWidth > 0) {
      vfParts.push(`scale=w=${maxWidth}:h=-2`);
    } else if (maxHeight && maxHeight > 0) {
      vfParts.push(`scale=w=-2:h=${maxHeight}`);
    }
    if (vfParts.length > 0) {
      args.push('-vf', vfParts.join(','));
    }

    args.push('-c:v', codecMap[format]);

    if (format !== 'png') {
      const clampedQuality = Math.max(1, Math.min(100, quality));
      if (format === 'jpg') {
        const qv = Math.round(2 + (100 - clampedQuality) * (29 / 99));
        args.push('-q:v', qv.toString());
      } else {
        args.push('-q:v', clampedQuality.toString());
      }
    }

    args.push('-y');

    return new Promise((resolve, reject) => {
      const outputPath = this.createTempOutputPath('screenshot', ext);
      const startedAt = this.nowMs();
      args.push(outputPath);

      this.logMediaDebug(
        `screenshot start ${inputDescription} timestamp=${timestamp} format=${format} maxWidth=${maxWidth ?? 'none'} maxHeight=${maxHeight ?? 'none'}`,
      );
      execFile('ffmpeg', args, { timeout: 30000 }, (error) => {
        if (error) {
          this.logMediaDebug(
            `screenshot failed ${inputDescription} elapsedMs=${this.elapsedMs(startedAt)} ${describeFfmpegFailureForDebugLog(error)}`,
          );
          reject(this.ffmpegError('screenshot generation', error));
          return;
        }

        try {
          const data = fs.readFileSync(outputPath);
          fs.unlinkSync(outputPath);
          this.logMediaDebug(
            `screenshot complete ${inputDescription} elapsedMs=${this.elapsedMs(startedAt)} bytes=${data.byteLength}`,
          );
          resolve(data);
        } catch (err) {
          reject(err);
        }
      });
    });
  }

  /**
   * Generate a small PNG icon suitable for desktop notifications.
   * Always outputs PNG format (known-good for Electron + Linux notification daemons).
   * Scaled to 256px width for fast encoding and small file size.
   */
  async generateNotificationIcon(videoPath: string, timestamp: number): Promise<Buffer> {
    return new Promise((resolve, reject) => {
      const outputPath = this.createTempOutputPath('notify_icon', 'png');

      execFile(
        'ffmpeg',
        [
          '-ss',
          timestamp.toString(),
          '-i',
          videoPath,
          '-vframes',
          '1',
          '-vf',
          'scale=256:256:force_original_aspect_ratio=decrease,pad=256:256:(ow-iw)/2:(oh-ih)/2:black',
          '-c:v',
          'png',
          '-y',
          outputPath,
        ],
        { timeout: 30000 },
        (error) => {
          if (error) {
            reject(this.ffmpegError('notification icon generation', error));
            return;
          }

          try {
            const data = fs.readFileSync(outputPath);
            fs.unlinkSync(outputPath);
            resolve(data);
          } catch (err) {
            reject(err);
          }
        },
      );
    });
  }

  async generateAnimatedImage(
    videoPath: MediaInput,
    startTime: number,
    endTime: number,
    padding: number = 0,
    options: {
      fps?: number;
      maxWidth?: number;
      maxHeight?: number;
      crf?: number;
      leadingStillDuration?: number;
    } = {},
  ): Promise<Buffer> {
    const { fps = 10, maxWidth = 640, maxHeight, crf = 35, leadingStillDuration = 0 } = options;
    const clampedFps = normalizeAnimatedImageFps(fps);
    const safePadding = Number.isFinite(padding) ? Math.max(0, padding) : 0;
    const start = Math.max(0, startTime - safePadding);
    const duration = roundDurationUpToNextFrameBoundary(endTime - start + safePadding, clampedFps);
    const totalLeadingStillDuration = Math.max(0, leadingStillDuration);
    const inputDescription = describeMediaInputForDebugLog(videoPath);

    const clampedCrf = Math.max(0, Math.min(63, crf));

    const encoderDetectionStartedAt = this.nowMs();
    const av1Encoder = await this.detectAv1Encoder();
    this.logMediaDebug(
      `animated-image encoder ${inputDescription} elapsedMs=${this.elapsedMs(encoderDetectionStartedAt)} encoder=${av1Encoder ?? 'none'}`,
    );
    if (!av1Encoder) {
      throw new Error(
        'No supported AV1 encoder found for animated AVIF (tried libaom-av1, libsvtav1, librav1e).',
      );
    }

    return new Promise((resolve, reject) => {
      const outputPath = this.createTempOutputPath('animation', 'avif');
      const mediaInput = normalizeMediaInput(videoPath);
      const startedAt = this.nowMs();

      const encoderArgs: string[] = ['-c:v', av1Encoder];
      if (av1Encoder === 'libaom-av1') {
        encoderArgs.push('-crf', clampedCrf.toString(), '-b:v', '0', '-cpu-used', '8');
      } else if (av1Encoder === 'libsvtav1') {
        encoderArgs.push('-crf', clampedCrf.toString(), '-preset', '8');
      } else {
        // librav1e
        encoderArgs.push('-qp', clampedCrf.toString(), '-speed', '8');
      }

      this.logMediaDebug(
        `animated-image start ${inputDescription} start=${start} duration=${duration} padding=${safePadding} fps=${clampedFps} maxWidth=${maxWidth ?? 'none'} maxHeight=${maxHeight ?? 'none'} crf=${clampedCrf} encoder=${av1Encoder}`,
      );
      execFile(
        'ffmpeg',
        [
          '-ss',
          start.toString(),
          '-t',
          duration.toString(),
          ...mediaInput.inputArgs,
          '-i',
          mediaInput.path,
          '-vf',
          buildAnimatedImageVideoFilter({
            fps: clampedFps,
            maxWidth,
            maxHeight,
            leadingStillDuration: totalLeadingStillDuration,
          }),
          ...encoderArgs,
          '-y',
          outputPath,
        ],
        { timeout: 60000 },
        (error) => {
          if (error) {
            this.logMediaDebug(
              `animated-image failed ${inputDescription} elapsedMs=${this.elapsedMs(startedAt)} ${describeFfmpegFailureForDebugLog(error)}`,
            );
            reject(this.ffmpegError('animation generation', error));
            return;
          }

          try {
            const data = fs.readFileSync(outputPath);
            fs.unlinkSync(outputPath);
            this.logMediaDebug(
              `animated-image complete ${inputDescription} elapsedMs=${this.elapsedMs(startedAt)} bytes=${data.byteLength}`,
            );
            resolve(data);
          } catch (err) {
            reject(err);
          }
        },
      );
    });
  }

  cleanup(): void {
    try {
      if (fs.existsSync(this.tempDir)) {
        fs.rmSync(this.tempDir, { recursive: true, force: true });
      }
    } catch (err) {
      log.error('Failed to cleanup media generator temp directory:', err);
    }
  }
}

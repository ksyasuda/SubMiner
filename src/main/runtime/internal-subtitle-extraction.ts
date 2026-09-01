import * as fs from 'fs';
import { spawn } from 'node:child_process';
import * as os from 'os';
import * as path from 'path';

import { resolveSubtitleSourcePath } from './subtitle-prefetch-source';
import { codecToExtension } from '../../subsync/utils';

export async function loadSubtitleSourceText(source: string): Promise<string> {
  if (/^https?:\/\//i.test(source)) {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 4000);
    try {
      const response = await fetch(source, { signal: controller.signal });
      if (!response.ok) {
        throw new Error(`Failed to download subtitle source (${response.status})`);
      }
      return await response.text();
    } finally {
      clearTimeout(timeoutId);
    }
  }

  const filePath = resolveSubtitleSourcePath(source);
  return fs.promises.readFile(filePath, 'utf8');
}

export type MpvSubtitleTrackLike = {
  type?: unknown;
  id?: unknown;
  selected?: unknown;
  external?: unknown;
  codec?: unknown;
  'ff-index'?: unknown;
  'external-filename'?: unknown;
};

export type ExtractedInternalSubtitleTrack = {
  path: string;
  cleanup: () => Promise<void>;
};

export type InternalSubtitleTrackExtractor = (
  ffmpegPath: string,
  videoPath: string,
  track: MpvSubtitleTrackLike,
) => Promise<ExtractedInternalSubtitleTrack | null>;

// Subtitle packets are interleaved through the container, so extraction reads the
// entire file. Network mounts move ~100 MB/s on gigabit, so large Bluray remuxes
// need well over 30 seconds.
const DEFAULT_EXTRACTION_TIMEOUT_MS = 120_000;

export function parseTrackId(value: unknown): number | null {
  if (typeof value === 'number' && Number.isInteger(value) && value >= 0) {
    return value;
  }
  if (typeof value === 'string') {
    const parsed = Number(value.trim());
    return Number.isInteger(parsed) && parsed >= 0 ? parsed : null;
  }
  return null;
}

export function buildFfmpegSubtitleExtractionArgs(
  videoPath: string,
  ffIndex: number,
  outputPath: string,
): string[] {
  const outputFormat = path.extname(outputPath).slice(1);
  if (!outputFormat) {
    throw new Error(`outputPath must include a file extension for ffmpeg format: ${outputPath}`);
  }
  return [
    '-hide_banner',
    '-nostdin',
    '-y',
    '-loglevel',
    'error',
    '-an',
    '-vn',
    '-i',
    videoPath,
    '-map',
    `0:${ffIndex}`,
    '-f',
    outputFormat,
    outputPath,
  ];
}

export async function extractInternalSubtitleTrackToTempFile(
  ffmpegPath: string,
  videoPath: string,
  track: MpvSubtitleTrackLike,
  options: { extractionTimeoutMs?: number; spawnArgsOverride?: string[] } = {},
): Promise<ExtractedInternalSubtitleTrack | null> {
  const ffIndex = parseTrackId(track['ff-index']);
  const codec = typeof track.codec === 'string' ? track.codec : null;
  const extension = codecToExtension(codec ?? undefined);
  if (ffIndex === null || extension === null) {
    return null;
  }

  const tempDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'subminer-sidebar-'));
  const outputPath = path.join(tempDir, `track_${ffIndex}.${extension}`);

  try {
    await new Promise<void>((resolve, reject) => {
      let settled = false;
      const child = spawn(
        ffmpegPath,
        options.spawnArgsOverride ??
          buildFfmpegSubtitleExtractionArgs(videoPath, ffIndex, outputPath),
      );
      const extractionTimeoutMs = options.extractionTimeoutMs ?? DEFAULT_EXTRACTION_TIMEOUT_MS;
      const timeoutId = setTimeout(() => {
        if (settled) {
          return;
        }
        settled = true;
        child.kill('SIGKILL');
        reject(new Error(`ffmpeg extraction timed out after ${extractionTimeoutMs}ms`));
      }, extractionTimeoutMs);
      const settle = (callback: () => void): void => {
        if (settled) {
          return;
        }
        settled = true;
        clearTimeout(timeoutId);
        callback();
      };
      let stderr = '';
      child.stderr.on('data', (chunk: Buffer) => {
        stderr += chunk.toString();
      });
      child.on('error', (error) => {
        settle(() => reject(error));
      });
      child.on('close', (code) => {
        settle(() => {
          if (code === 0) {
            resolve();
            return;
          }
          reject(new Error(stderr.trim() || `ffmpeg exited with code ${code ?? 'unknown'}`));
        });
      });
    });
  } catch (error) {
    await fs.promises.rm(tempDir, { recursive: true, force: true }).catch(() => undefined);
    throw error;
  }

  return {
    path: outputPath,
    cleanup: async () => {
      await fs.promises.rm(tempDir, { recursive: true, force: true });
    },
  };
}

type CachedExtraction = {
  promise: Promise<ExtractedInternalSubtitleTrack | null>;
};

function buildCachedExtractionKey(
  ffmpegPath: string,
  videoPath: string,
  track: MpvSubtitleTrackLike,
): string {
  const codec = typeof track.codec === 'string' ? track.codec : null;
  return JSON.stringify([ffmpegPath, videoPath, parseTrackId(track['ff-index']), codec]);
}

const releaseCachedExtraction = async (): Promise<void> => {};

/**
 * Owns extracted subtitle files for the active media and shares one extraction between callers.
 * Caller cleanup releases only its view; clear removes the owned files on media changes or quit.
 */
export function createCachedInternalSubtitleTrackExtractor(
  deps: { extract?: InternalSubtitleTrackExtractor } = {},
): {
  extract: InternalSubtitleTrackExtractor;
  clear: () => void;
} {
  const extractTrack = deps.extract ?? extractInternalSubtitleTrackToTempFile;
  const extractions = new Map<string, CachedExtraction>();

  const extract: InternalSubtitleTrackExtractor = async (ffmpegPath, videoPath, track) => {
    const key = buildCachedExtractionKey(ffmpegPath, videoPath, track);
    let cached = extractions.get(key);
    if (!cached) {
      const next: CachedExtraction = {
        promise: extractTrack(ffmpegPath, videoPath, track),
      };
      cached = next;
      extractions.set(key, next);
      void next.promise.catch(() => {
        if (extractions.get(key) === next) {
          extractions.delete(key);
        }
      });
    }

    const result = await cached.promise;
    if (extractions.get(key) !== cached || !result) {
      return null;
    }

    return {
      path: result.path,
      cleanup: releaseCachedExtraction,
    };
  };

  const clear = (): void => {
    const staleExtractions = [...extractions.values()];
    extractions.clear();
    for (const extraction of staleExtractions) {
      void extraction.promise.then((result) => result?.cleanup()).catch(() => undefined);
    }
  };

  return { extract, clear };
}

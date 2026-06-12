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

export function parseTrackId(value: unknown): number | null {
  if (typeof value === 'number' && Number.isInteger(value)) {
    return value;
  }
  if (typeof value === 'string') {
    const parsed = Number(value.trim());
    return Number.isInteger(parsed) ? parsed : null;
  }
  return null;
}

export function buildFfmpegSubtitleExtractionArgs(
  videoPath: string,
  ffIndex: number,
  outputPath: string,
): string[] {
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
    path.extname(outputPath).slice(1),
    outputPath,
  ];
}

export async function extractInternalSubtitleTrackToTempFile(
  ffmpegPath: string,
  videoPath: string,
  track: MpvSubtitleTrackLike,
): Promise<{ path: string; cleanup: () => Promise<void> } | null> {
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
      const child = spawn(
        ffmpegPath,
        buildFfmpegSubtitleExtractionArgs(videoPath, ffIndex, outputPath),
      );
      let stderr = '';
      child.stderr.on('data', (chunk: Buffer) => {
        stderr += chunk.toString();
      });
      child.on('error', (error) => {
        reject(error);
      });
      child.on('close', (code) => {
        if (code === 0) {
          resolve();
          return;
        }
        reject(new Error(stderr.trim() || `ffmpeg exited with code ${code ?? 'unknown'}`));
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

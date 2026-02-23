import crypto from 'node:crypto';
import { spawn as nodeSpawn } from 'node:child_process';
import * as fs from 'node:fs';
import {
  deriveCanonicalTitle,
  emptyMetadata,
  hashToCode,
  parseFps,
  toNullableInt,
} from './reducer';
import { SOURCE_TYPE_LOCAL, type ProbeMetadata, type VideoMetadata } from './types';

type SpawnFn = typeof nodeSpawn;

interface FsDeps {
  createReadStream: typeof fs.createReadStream;
  promises: {
    stat: typeof fs.promises.stat;
  };
}

interface MetadataDeps {
  spawn?: SpawnFn;
  fs?: FsDeps;
}

export async function computeSha256(
  mediaPath: string,
  deps: MetadataDeps = {},
): Promise<string | null> {
  const fileSystem = deps.fs ?? fs;
  return new Promise((resolve) => {
    const file = fileSystem.createReadStream(mediaPath);
    const digest = crypto.createHash('sha256');
    file.on('data', (chunk) => digest.update(chunk));
    file.on('end', () => resolve(digest.digest('hex')));
    file.on('error', () => resolve(null));
  });
}

export function runFfprobe(mediaPath: string, deps: MetadataDeps = {}): Promise<ProbeMetadata> {
  const spawn = deps.spawn ?? nodeSpawn;
  return new Promise((resolve) => {
    const child = spawn('ffprobe', [
      '-v',
      'error',
      '-print_format',
      'json',
      '-show_entries',
      'stream=codec_type,codec_tag_string,width,height,avg_frame_rate,bit_rate',
      '-show_entries',
      'format=duration,bit_rate',
      mediaPath,
    ]);

    let output = '';
    let errorOutput = '';
    child.stdout.on('data', (chunk) => {
      output += chunk.toString('utf-8');
    });
    child.stderr.on('data', (chunk) => {
      errorOutput += chunk.toString('utf-8');
    });
    child.on('error', () => resolve(emptyMetadata()));
    child.on('close', () => {
      if (errorOutput && output.length === 0) {
        resolve(emptyMetadata());
        return;
      }

      try {
        const parsed = JSON.parse(output) as {
          format?: { duration?: string; bit_rate?: string };
          streams?: Array<{
            codec_type?: string;
            codec_tag_string?: string;
            width?: number;
            height?: number;
            avg_frame_rate?: string;
            bit_rate?: string;
          }>;
        };

        const durationText = parsed.format?.duration;
        const bitrateText = parsed.format?.bit_rate;
        const durationMs = Number(durationText) ? Math.round(Number(durationText) * 1000) : null;
        const bitrateKbps = Number(bitrateText) ? Math.round(Number(bitrateText) / 1000) : null;

        let codecId: number | null = null;
        let containerId: number | null = null;
        let widthPx: number | null = null;
        let heightPx: number | null = null;
        let fpsX100: number | null = null;
        let audioCodecId: number | null = null;

        for (const stream of parsed.streams ?? []) {
          if (stream.codec_type === 'video') {
            widthPx = toNullableInt(stream.width);
            heightPx = toNullableInt(stream.height);
            fpsX100 = parseFps(stream.avg_frame_rate);
            codecId = hashToCode(stream.codec_tag_string);
            containerId = 0;
          }
          if (stream.codec_type === 'audio') {
            audioCodecId = hashToCode(stream.codec_tag_string);
            if (audioCodecId && audioCodecId > 0) {
              break;
            }
          }
        }

        resolve({
          durationMs,
          codecId,
          containerId,
          widthPx,
          heightPx,
          fpsX100,
          bitrateKbps,
          audioCodecId,
        });
      } catch {
        resolve(emptyMetadata());
      }
    });
  });
}

export async function getLocalVideoMetadata(
  mediaPath: string,
  deps: MetadataDeps = {},
): Promise<VideoMetadata> {
  const fileSystem = deps.fs ?? fs;
  const hash = await computeSha256(mediaPath, deps);
  const info = await runFfprobe(mediaPath, deps);
  const stat = await fileSystem.promises.stat(mediaPath);
  return {
    sourceType: SOURCE_TYPE_LOCAL,
    canonicalTitle: deriveCanonicalTitle(mediaPath),
    durationMs: info.durationMs || 0,
    fileSizeBytes: Number.isFinite(stat.size) ? stat.size : null,
    codecId: info.codecId ?? null,
    containerId: info.containerId ?? null,
    widthPx: info.widthPx ?? null,
    heightPx: info.heightPx ?? null,
    fpsX100: info.fpsX100 ?? null,
    bitrateKbps: info.bitrateKbps ?? null,
    audioCodecId: info.audioCodecId ?? null,
    hashSha256: hash,
    screenshotPath: null,
    metadataJson: null,
  };
}

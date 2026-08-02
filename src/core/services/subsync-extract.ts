import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
  codecToExtension,
  fileExists,
  MpvTrack,
  runCommand,
  summarizeCommandFailure,
} from '../../subsync/utils';
import { downloadToFile, isRemoteMediaPath } from '../../jimaku/utils';
import { createLogger } from '../../logger';
import {
  ResolvedMpvHttpHeaders,
  toFfmpegInputHttpArgs,
  toRequestHeaders,
} from './mpv-http-headers';

const logger = createLogger('subsync');

export interface FileExtractionResult {
  path: string;
  temporary: boolean;
}

export interface SubtitleExtractionInput {
  ffmpegPath: string;
  videoPath: string;
  track: MpvTrack;
  /** mpv's request context, so stream-hosted tracks stay reachable. */
  httpHeaders: ResolvedMpvHttpHeaders | null;
}

function extensionForUrl(url: string, fallback: string): string {
  try {
    const parsed = new URL(url);
    const ext = path.extname(parsed.pathname).replace(/^\./, '').toLowerCase();
    return /^[a-z0-9]{1,5}$/.test(ext) ? ext : fallback;
  } catch {
    return fallback;
  }
}

/**
 * Pull a subtitle track mpv loaded from a URL down to disk.
 *
 * Extension-backed and Jellyfin streams add their subtitles by URL, and neither
 * alass nor ffmpeg can read one — the whole subsync path used to reject these
 * with "Subtitle file not found: https://…".
 */
async function downloadRemoteSubtitleTrack(
  url: string,
  httpHeaders: ResolvedMpvHttpHeaders | null,
): Promise<FileExtractionResult> {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-subsync-'));
  const outputPath = path.join(tempDir, `remote_track.${extensionForUrl(url, 'srt')}`);

  const result = await downloadToFile(url, outputPath, toRequestHeaders(httpHeaders));
  if (!result.ok) {
    throw new Error(`Failed to download subtitle track: ${result.error?.error ?? 'unknown error'}`);
  }
  if (!fileExists(outputPath)) {
    throw new Error(`Downloaded subtitle track is missing: ${url}`);
  }

  logger.info(`Downloaded remote subtitle track to ${outputPath}`);
  return { path: outputPath, temporary: true };
}

async function extractInternalTrack(input: SubtitleExtractionInput): Promise<FileExtractionResult> {
  const ffIndex = input.track['ff-index'];
  const extension = codecToExtension(input.track.codec);
  if (typeof ffIndex !== 'number' || !Number.isInteger(ffIndex) || ffIndex < 0) {
    throw new Error('Internal subtitle track has no valid ff-index');
  }
  if (!extension) {
    throw new Error(`Unsupported subtitle codec: ${input.track.codec ?? 'unknown'}`);
  }

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-subsync-'));
  const outputPath = path.join(tempDir, `track_${ffIndex}.${extension}`);
  // Header args configure the HTTP demuxer, so they belong before `-i`.
  const httpArgs = isRemoteMediaPath(input.videoPath)
    ? toFfmpegInputHttpArgs(input.httpHeaders)
    : [];

  const extraction = await runCommand(input.ffmpegPath, [
    '-hide_banner',
    '-nostdin',
    '-y',
    '-loglevel',
    'error',
    ...httpArgs,
    '-an',
    '-vn',
    '-i',
    input.videoPath,
    '-map',
    `0:${ffIndex}`,
    '-f',
    extension,
    outputPath,
  ]);

  if (!extraction.ok || !fileExists(outputPath)) {
    throw new Error(
      `Failed to extract internal subtitle track with ffmpeg: ${summarizeCommandFailure(
        'ffmpeg',
        extraction,
      )}`,
    );
  }

  return { path: outputPath, temporary: true };
}

export async function extractSubtitleTrackToFile(
  input: SubtitleExtractionInput,
): Promise<FileExtractionResult> {
  if (input.track.external) {
    const externalPath = input.track['external-filename'];
    if (typeof externalPath !== 'string' || externalPath.length === 0) {
      throw new Error('External subtitle track has no file path');
    }
    if (isRemoteMediaPath(externalPath)) {
      return downloadRemoteSubtitleTrack(externalPath, input.httpHeaders);
    }
    if (!fileExists(externalPath)) {
      throw new Error(`Subtitle file not found: ${externalPath}`);
    }
    return { path: externalPath, temporary: false };
  }

  return extractInternalTrack(input);
}

/**
 * Drop a temporary extraction once alass/ffsubsync is done with it.
 *
 * `preservePath` is the retimed subtitle mpv is about to load. With
 * `replace: false` it lands in this same temp directory, so the removal has to
 * skip it — and the directory removal must stay non-recursive, since that is
 * what keeps the retimed file alive for mpv.
 */
export function cleanupTemporaryFile(
  extraction: FileExtractionResult,
  preservePath?: string,
): void {
  if (!extraction.temporary) return;
  if (preservePath && path.resolve(preservePath) === path.resolve(extraction.path)) return;

  try {
    if (fileExists(extraction.path)) {
      fs.unlinkSync(extraction.path);
    }
  } catch {}
  try {
    const dir = path.dirname(extraction.path);
    if (fs.existsSync(dir)) {
      fs.rmdirSync(dir);
    }
  } catch {}
}

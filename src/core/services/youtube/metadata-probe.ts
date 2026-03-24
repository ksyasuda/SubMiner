import { spawn } from 'node:child_process';
import type { YoutubeVideoMetadata } from '../immersion-tracker/types';

const YOUTUBE_METADATA_PROBE_TIMEOUT_MS = 15_000;

type YtDlpThumbnail = {
  url?: string;
  width?: number;
  height?: number;
};

type YtDlpYoutubeMetadata = {
  id?: string;
  title?: string;
  webpage_url?: string;
  thumbnail?: string;
  thumbnails?: YtDlpThumbnail[];
  channel_id?: string;
  channel?: string;
  channel_url?: string;
  uploader_id?: string;
  uploader_url?: string;
  description?: string;
};

function runCapture(
  command: string,
  args: string[],
  timeoutMs = YOUTUBE_METADATA_PROBE_TIMEOUT_MS,
): Promise<{ stdout: string; stderr: string }> {
  return new Promise((resolve, reject) => {
    const proc = spawn(command, args, { stdio: ['ignore', 'pipe', 'pipe'] });
    let stdout = '';
    let stderr = '';
    const timer = setTimeout(() => {
      proc.kill();
      reject(new Error(`yt-dlp timed out after ${timeoutMs}ms`));
    }, timeoutMs);
    proc.stdout.setEncoding('utf8');
    proc.stderr.setEncoding('utf8');
    proc.stdout.on('data', (chunk) => {
      stdout += String(chunk);
    });
    proc.stderr.on('data', (chunk) => {
      stderr += String(chunk);
    });
    proc.once('error', (error) => {
      clearTimeout(timer);
      reject(error);
    });
    proc.once('close', (code) => {
      clearTimeout(timer);
      if (code === 0) {
        resolve({ stdout, stderr });
        return;
      }
      reject(new Error(stderr.trim() || `yt-dlp exited with status ${code ?? 'unknown'}`));
    });
  });
}

function pickChannelThumbnail(thumbnails: YtDlpThumbnail[] | undefined): string | null {
  if (!Array.isArray(thumbnails)) return null;
  for (const thumbnail of thumbnails) {
    const candidate = thumbnail.url?.trim();
    if (!candidate) continue;
    if (candidate.includes('/vi/')) continue;
    if (
      typeof thumbnail.width === 'number' &&
      typeof thumbnail.height === 'number' &&
      thumbnail.width > 0 &&
      thumbnail.height > 0
    ) {
      const ratio = thumbnail.width / thumbnail.height;
      if (ratio >= 0.8 && ratio <= 1.25) {
        return candidate;
      }
      continue;
    }
    if (candidate.includes('yt3.googleusercontent.com')) {
      return candidate;
    }
  }
  return null;
}

export async function probeYoutubeVideoMetadata(
  targetUrl: string,
): Promise<YoutubeVideoMetadata | null> {
  const { stdout } = await runCapture('yt-dlp', [
    '--dump-single-json',
    '--no-warnings',
    '--skip-download',
    targetUrl,
  ]);
  let info: YtDlpYoutubeMetadata;
  try {
    info = JSON.parse(stdout) as YtDlpYoutubeMetadata;
  } catch {
    return null;
  }
  const youtubeVideoId = info.id?.trim();
  const videoUrl = info.webpage_url?.trim() || targetUrl.trim();
  if (!youtubeVideoId || !videoUrl) {
    return null;
  }

  return {
    youtubeVideoId,
    videoUrl,
    videoTitle: info.title?.trim() || null,
    videoThumbnailUrl: info.thumbnail?.trim() || null,
    channelId: info.channel_id?.trim() || null,
    channelName: info.channel?.trim() || null,
    channelUrl: info.channel_url?.trim() || null,
    channelThumbnailUrl: pickChannelThumbnail(info.thumbnails),
    uploaderId: info.uploader_id?.trim() || null,
    uploaderUrl: info.uploader_url?.trim() || null,
    description: info.description?.trim() || null,
    metadataJson: JSON.stringify(info),
  };
}

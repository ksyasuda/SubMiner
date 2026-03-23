import { spawn } from 'node:child_process';
import type { YoutubeTrackOption } from '../../../types';
import {
  formatYoutubeTrackLabel,
  normalizeYoutubeLangCode,
  type YoutubeTrackKind,
} from './labels';

export type YoutubeTrackProbeResult = {
  videoId: string;
  title: string;
  tracks: YoutubeTrackOption[];
};

type YtDlpSubtitleEntry = Array<{ ext?: string; name?: string; url?: string }>;

type YtDlpInfo = {
  id?: string;
  title?: string;
  subtitles?: Record<string, YtDlpSubtitleEntry>;
  automatic_captions?: Record<string, YtDlpSubtitleEntry>;
};

function runCapture(command: string, args: string[]): Promise<{ stdout: string; stderr: string }> {
  return new Promise((resolve, reject) => {
    const proc = spawn(command, args, { stdio: ['ignore', 'pipe', 'pipe'] });
    let stdout = '';
    let stderr = '';
    proc.stdout.setEncoding('utf8');
    proc.stderr.setEncoding('utf8');
    proc.stdout.on('data', (chunk) => {
      stdout += String(chunk);
    });
    proc.stderr.on('data', (chunk) => {
      stderr += String(chunk);
    });
    proc.once('error', reject);
    proc.once('close', (code) => {
      if (code === 0) {
        resolve({ stdout, stderr });
        return;
      }
      reject(new Error(stderr.trim() || `yt-dlp exited with status ${code ?? 'unknown'}`));
    });
  });
}

function choosePreferredFormat(
  formats: YtDlpSubtitleEntry,
  kind: YoutubeTrackKind,
): { ext: string; url: string; title?: string } | null {
  const preferredOrder =
    kind === 'auto'
      ? ['srv3', 'srv2', 'srv1', 'vtt', 'srt', 'ttml', 'json3']
      : ['srt', 'vtt', 'srv3', 'srv2', 'srv1', 'ttml', 'json3'];
  for (const ext of preferredOrder) {
    const match = formats.find(
      (format) => typeof format.url === 'string' && format.url && format.ext === ext,
    );
    if (match?.url) {
      return { ext, url: match.url, title: match.name?.trim() || undefined };
    }
  }

  const fallback = formats.find((format) => typeof format.url === 'string' && format.url);
  if (!fallback?.url) {
    return null;
  }

  return {
    ext: fallback.ext?.trim() || 'vtt',
    url: fallback.url,
    title: fallback.name?.trim() || undefined,
  };
}

function toTracks(entries: Record<string, YtDlpSubtitleEntry> | undefined, kind: YoutubeTrackKind) {
  const tracks: YoutubeTrackOption[] = [];
  if (!entries) return tracks;
  for (const [language, formats] of Object.entries(entries)) {
    if (!Array.isArray(formats) || formats.length === 0) continue;
    const preferredFormat = choosePreferredFormat(formats, kind);
    if (!preferredFormat) continue;
    const sourceLanguage = language.trim() || language;
    const normalizedLanguage = normalizeYoutubeLangCode(sourceLanguage) || sourceLanguage;
    const title = preferredFormat.title;
    tracks.push({
      id: `${kind}:${sourceLanguage}`,
      language: normalizedLanguage,
      sourceLanguage,
      kind,
      title,
      label: formatYoutubeTrackLabel({ language: normalizedLanguage, kind, title }),
      downloadUrl: preferredFormat.url,
      fileExtension: preferredFormat.ext,
    });
  }
  return tracks;
}

export type { YoutubeTrackOption };

export async function probeYoutubeTracks(targetUrl: string): Promise<YoutubeTrackProbeResult> {
  const { stdout } = await runCapture('yt-dlp', ['--dump-single-json', '--no-warnings', targetUrl]);
  const info = JSON.parse(stdout) as YtDlpInfo;
  const tracks = [...toTracks(info.subtitles, 'manual'), ...toTracks(info.automatic_captions, 'auto')];
  return {
    videoId: info.id || '',
    title: info.title || '',
    tracks,
  };
}

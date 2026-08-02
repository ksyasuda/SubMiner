import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import type { YoutubeTrackOption } from './track-probe';
import { getYoutubeYtDlpCommand, YTDLP_SINGLE_VIDEO_ARG } from './ytdlp-command';
import {
  convertYoutubeTimedTextToVtt,
  isYoutubeTimedTextExtension,
  normalizeYoutubeAutoVtt,
} from './timedtext';

const YOUTUBE_SUBTITLE_EXTENSIONS = new Set(['.srt', '.vtt', '.ass']);
const YOUTUBE_BATCH_PREFIX = 'youtube-batch';
const YOUTUBE_DOWNLOAD_TIMEOUT_MS = 15_000;

function sanitizeFilenameSegment(value: string): string {
  const sanitized = value
    .trim()
    .replace(/[^a-z0-9_-]+/gi, '-')
    .replace(/-+/g, '-');
  return sanitized.replace(/^-+|-+$/g, '') || 'unknown';
}

function createFetchTimeoutSignal(timeoutMs: number): AbortSignal | undefined {
  if (typeof AbortSignal !== 'undefined' && typeof AbortSignal.timeout === 'function') {
    return AbortSignal.timeout(timeoutMs);
  }
  return undefined;
}

function runCapture(
  command: string,
  args: string[],
  timeoutMs = YOUTUBE_DOWNLOAD_TIMEOUT_MS,
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

function runCaptureDetailed(
  command: string,
  args: string[],
  timeoutMs = YOUTUBE_DOWNLOAD_TIMEOUT_MS,
): Promise<{ stdout: string; stderr: string; code: number }> {
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
      resolve({ stdout, stderr, code: code ?? 1 });
    });
  });
}

function pickLatestSubtitleFile(dir: string, prefix: string): string | null {
  const entries = fs.readdirSync(dir).map((name) => path.join(dir, name));
  const candidates = entries.filter((candidate) => {
    const basename = path.basename(candidate);
    const ext = path.extname(basename).toLowerCase();
    return basename.startsWith(prefix) && YOUTUBE_SUBTITLE_EXTENSIONS.has(ext);
  });
  candidates.sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs);
  return candidates[0] ?? null;
}

function pickLatestSubtitleFileForLanguage(
  dir: string,
  prefix: string,
  sourceLanguage: string,
): string | null {
  const entries = fs.readdirSync(dir).map((name) => path.join(dir, name));
  const candidates = entries.filter((candidate) => {
    const basename = path.basename(candidate);
    const ext = path.extname(basename).toLowerCase();
    return (
      basename.startsWith(`${prefix}.`) &&
      basename.includes(`.${sourceLanguage}.`) &&
      YOUTUBE_SUBTITLE_EXTENSIONS.has(ext)
    );
  });
  candidates.sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs);
  return candidates[0] ?? null;
}

export function buildDownloadArgs(input: {
  targetUrl: string;
  outputTemplate: string;
  sourceLanguages: string[];
  includeAutoSubs: boolean;
  includeManualSubs: boolean;
}): string[] {
  const args = [YTDLP_SINGLE_VIDEO_ARG, '--skip-download', '--no-warnings'];
  if (input.includeAutoSubs) {
    args.push('--write-auto-subs');
  }
  if (input.includeManualSubs) {
    args.push('--write-subs');
  }
  args.push(
    '--sub-format',
    'srt/vtt/best',
    '--sub-langs',
    input.sourceLanguages.join(','),
    '-o',
    input.outputTemplate,
    input.targetUrl,
  );
  return args;
}

async function downloadSubtitleFromUrl(input: {
  outputDir: string;
  prefix: string;
  track: YoutubeTrackOption;
}): Promise<{ path: string }> {
  if (!input.track.downloadUrl) {
    throw new Error(`No direct subtitle URL available for ${input.track.sourceLanguage}`);
  }
  const ext = (input.track.fileExtension?.trim().toLowerCase() || 'vtt').replace(/[^a-z0-9]+/g, '');
  const safeExt = isYoutubeTimedTextExtension(ext)
    ? 'vtt'
    : YOUTUBE_SUBTITLE_EXTENSIONS.has(`.${ext}`)
      ? ext
      : 'vtt';
  const safeSourceLanguage = sanitizeFilenameSegment(input.track.sourceLanguage);
  const targetPath = path.join(input.outputDir, `${input.prefix}.${safeSourceLanguage}.${safeExt}`);
  const response = await fetch(input.track.downloadUrl, {
    signal: createFetchTimeoutSignal(YOUTUBE_DOWNLOAD_TIMEOUT_MS),
  });
  if (!response.ok) {
    throw new Error(`HTTP ${response.status} while downloading ${input.track.sourceLanguage}`);
  }
  const body = await response.text();
  const normalizedBody = isYoutubeTimedTextExtension(ext)
    ? convertYoutubeTimedTextToVtt(body)
    : input.track.kind === 'auto' && safeExt === 'vtt'
      ? normalizeYoutubeAutoVtt(body)
      : body;
  fs.writeFileSync(targetPath, normalizedBody, 'utf8');
  return { path: targetPath };
}

function canDownloadSubtitleFromUrl(track: YoutubeTrackOption): boolean {
  if (!track.downloadUrl) {
    return false;
  }

  const ext = (track.fileExtension?.trim().toLowerCase() || 'vtt').replace(/[^a-z0-9]+/g, '');
  return isYoutubeTimedTextExtension(ext) || YOUTUBE_SUBTITLE_EXTENSIONS.has(`.${ext}`);
}

function normalizeDownloadedAutoSubtitle(pathname: string, track: YoutubeTrackOption): void {
  if (track.kind !== 'auto' || path.extname(pathname).toLowerCase() !== '.vtt') {
    return;
  }
  const content = fs.readFileSync(pathname, 'utf8');
  const normalized = normalizeYoutubeAutoVtt(content);
  if (normalized !== content) {
    fs.writeFileSync(pathname, normalized, 'utf8');
  }
}

export async function downloadYoutubeSubtitleTrack(input: {
  targetUrl: string;
  outputDir: string;
  track: YoutubeTrackOption;
}): Promise<{ path: string }> {
  fs.mkdirSync(input.outputDir, { recursive: true });
  const prefix = input.track.id.replace(/[^a-z0-9_-]+/gi, '-');
  for (const name of fs.readdirSync(input.outputDir)) {
    if (name.startsWith(prefix)) {
      try {
        fs.rmSync(path.join(input.outputDir, name), { force: true });
      } catch {
        // ignore stale files
      }
    }
  }
  if (canDownloadSubtitleFromUrl(input.track)) {
    return await downloadSubtitleFromUrl({
      outputDir: input.outputDir,
      prefix,
      track: input.track,
    });
  }
  const outputTemplate = path.join(input.outputDir, `${prefix}.%(ext)s`);
  const args = [
    ...buildDownloadArgs({
      targetUrl: input.targetUrl,
      outputTemplate,
      sourceLanguages: [input.track.sourceLanguage],
      includeAutoSubs: input.track.kind === 'auto',
      includeManualSubs: input.track.kind === 'manual',
    }),
  ];

  await runCapture(getYoutubeYtDlpCommand(), args);
  const subtitlePath = pickLatestSubtitleFile(input.outputDir, prefix);
  if (!subtitlePath) {
    throw new Error(`No subtitle file was downloaded for ${input.track.sourceLanguage}`);
  }
  normalizeDownloadedAutoSubtitle(subtitlePath, input.track);
  return { path: subtitlePath };
}

export async function downloadYoutubeSubtitleTracks(input: {
  targetUrl: string;
  outputDir: string;
  tracks: YoutubeTrackOption[];
}): Promise<Map<string, string>> {
  fs.mkdirSync(input.outputDir, { recursive: true });
  const hasDuplicateSourceLanguages =
    new Set(input.tracks.map((track) => track.sourceLanguage)).size !== input.tracks.length;
  for (const name of fs.readdirSync(input.outputDir)) {
    if (name.startsWith(`${YOUTUBE_BATCH_PREFIX}.`)) {
      try {
        fs.rmSync(path.join(input.outputDir, name), { force: true });
      } catch {
        // ignore stale files
      }
    }
  }
  if (hasDuplicateSourceLanguages || input.tracks.every(canDownloadSubtitleFromUrl)) {
    const results = new Map<string, string>();
    for (const track of input.tracks) {
      const download = await downloadSubtitleFromUrl({
        outputDir: input.outputDir,
        prefix: track.id.replace(/[^a-z0-9_-]+/gi, '-'),
        track,
      });
      results.set(track.id, download.path);
    }
    return results;
  }

  const outputTemplate = path.join(input.outputDir, `${YOUTUBE_BATCH_PREFIX}.%(ext)s`);
  const includeAutoSubs = input.tracks.some((track) => track.kind === 'auto');
  const includeManualSubs = input.tracks.some((track) => track.kind === 'manual');

  const result = await runCaptureDetailed(
    getYoutubeYtDlpCommand(),
    buildDownloadArgs({
      targetUrl: input.targetUrl,
      outputTemplate,
      sourceLanguages: input.tracks.map((track) => track.sourceLanguage),
      includeAutoSubs,
      includeManualSubs,
    }),
  );

  const results = new Map<string, string>();
  for (const track of input.tracks) {
    const subtitlePath = pickLatestSubtitleFileForLanguage(
      input.outputDir,
      YOUTUBE_BATCH_PREFIX,
      track.sourceLanguage,
    );
    if (subtitlePath) {
      normalizeDownloadedAutoSubtitle(subtitlePath, track);
      results.set(track.id, subtitlePath);
    }
  }
  if (results.size > 0) {
    return results;
  }
  if (result.code !== 0) {
    throw new Error(result.stderr.trim() || `yt-dlp exited with status ${result.code}`);
  }
  throw new Error(
    `No subtitle file was downloaded for ${input.tracks.map((track) => track.sourceLanguage).join(',')}`,
  );
}

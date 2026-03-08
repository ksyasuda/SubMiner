import fs from 'node:fs';
import path from 'node:path';

import type { Args } from '../types.js';
import { YOUTUBE_AUDIO_EXTENSIONS } from '../types.js';
import { runExternalCommand } from '../util.js';

export function findAudioFile(tempDir: string, preferredExt: string): string | null {
  const entries = fs.readdirSync(tempDir);
  const audioFiles: Array<{ path: string; ext: string; mtimeMs: number }> = [];
  for (const name of entries) {
    const fullPath = path.join(tempDir, name);
    let stat: fs.Stats;
    try {
      stat = fs.statSync(fullPath);
    } catch {
      continue;
    }
    if (!stat.isFile()) continue;
    const ext = path.extname(name).toLowerCase();
    if (!YOUTUBE_AUDIO_EXTENSIONS.has(ext)) continue;
    audioFiles.push({ path: fullPath, ext, mtimeMs: stat.mtimeMs });
  }
  if (audioFiles.length === 0) return null;
  const preferred = audioFiles.find((entry) => entry.ext === `.${preferredExt.toLowerCase()}`);
  if (preferred) return preferred.path;
  audioFiles.sort((a, b) => b.mtimeMs - a.mtimeMs);
  return audioFiles[0]?.path ?? null;
}

export async function convertAudioForWhisper(inputPath: string, tempDir: string): Promise<string> {
  const wavPath = path.join(tempDir, 'whisper-input.wav');
  await runExternalCommand('ffmpeg', [
    '-y',
    '-loglevel',
    'error',
    '-i',
    inputPath,
    '-ar',
    '16000',
    '-ac',
    '1',
    '-c:a',
    'pcm_s16le',
    wavPath,
  ]);
  if (!fs.existsSync(wavPath)) {
    throw new Error(`Failed to prepare whisper audio input: ${wavPath}`);
  }
  return wavPath;
}

export async function downloadYoutubeAudio(
  target: string,
  args: Args,
  tempDir: string,
  childTracker?: Set<ReturnType<typeof import('node:child_process').spawn>>,
): Promise<string> {
  await runExternalCommand(
    'yt-dlp',
    [
      '-f',
      'bestaudio/best',
      '--extract-audio',
      '--audio-format',
      args.youtubeSubgenAudioFormat,
      '--no-warnings',
      '-o',
      path.join(tempDir, '%(id)s.%(ext)s'),
      target,
    ],
    {
      logLevel: args.logLevel,
      commandLabel: 'yt-dlp:audio',
      streamOutput: true,
    },
    childTracker,
  );
  const audioPath = findAudioFile(tempDir, args.youtubeSubgenAudioFormat);
  if (!audioPath) {
    throw new Error('Audio extraction succeeded, but no audio file was found.');
  }
  return audioPath;
}

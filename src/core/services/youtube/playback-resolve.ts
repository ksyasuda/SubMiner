import { spawn } from 'node:child_process';

const YOUTUBE_PLAYBACK_RESOLVE_TIMEOUT_MS = 15_000;
const DEFAULT_PLAYBACK_FORMAT = 'b';

function runCapture(
  command: string,
  args: string[],
  timeoutMs = YOUTUBE_PLAYBACK_RESOLVE_TIMEOUT_MS,
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

export async function resolveYoutubePlaybackUrl(
  targetUrl: string,
  format = DEFAULT_PLAYBACK_FORMAT,
): Promise<string> {
  const ytDlpCommand = process.env.SUBMINER_YTDLP_BIN?.trim() || 'yt-dlp';
  const { stdout } = await runCapture(ytDlpCommand, [
    '--get-url',
    '--no-warnings',
    '-f',
    format,
    targetUrl,
  ]);
  const playbackUrl =
    stdout
      .split(/\r?\n/)
      .map((line) => line.trim())
      .find((line) => line.length > 0) ?? '';
  if (!playbackUrl) {
    throw new Error('yt-dlp returned empty output while resolving YouTube playback URL');
  }
  return playbackUrl;
}

import { spawn } from 'node:child_process';
import { getYoutubeYtDlpCommand, YTDLP_SINGLE_VIDEO_ARG } from './ytdlp-command';

const YOUTUBE_PLAYBACK_RESOLVE_TIMEOUT_MS = 15_000;
const DEFAULT_PLAYBACK_FORMAT = 'b';
const MAX_CAPTURE_BYTES = 1024 * 1024;

function terminateCaptureProcess(proc: ReturnType<typeof spawn>): void {
  if (proc.killed) {
    return;
  }
  try {
    proc.kill('SIGKILL');
  } catch {
    proc.kill();
  }
}

function runCapture(
  command: string,
  args: string[],
  timeoutMs = YOUTUBE_PLAYBACK_RESOLVE_TIMEOUT_MS,
): Promise<{ stdout: string; stderr: string }> {
  return new Promise((resolve, reject) => {
    const proc = spawn(command, args, { stdio: ['ignore', 'pipe', 'pipe'] });
    let stdout = '';
    let stderr = '';
    let settled = false;
    const cleanup = (): void => {
      clearTimeout(timer);
      proc.stdout.removeAllListeners('data');
      proc.stderr.removeAllListeners('data');
      proc.removeAllListeners('error');
      proc.removeAllListeners('close');
    };
    const rejectOnce = (error: Error): void => {
      if (settled) return;
      settled = true;
      cleanup();
      reject(error);
    };
    const resolveOnce = (result: { stdout: string; stderr: string }): void => {
      if (settled) return;
      settled = true;
      cleanup();
      resolve(result);
    };
    const appendChunk = (
      current: string,
      chunk: unknown,
      streamName: 'stdout' | 'stderr',
    ): string => {
      const next = current + String(chunk);
      if (Buffer.byteLength(next, 'utf8') > MAX_CAPTURE_BYTES) {
        terminateCaptureProcess(proc);
        rejectOnce(new Error(`yt-dlp ${streamName} exceeded ${MAX_CAPTURE_BYTES} bytes`));
      }
      return next;
    };
    const timer = setTimeout(() => {
      terminateCaptureProcess(proc);
      rejectOnce(new Error(`yt-dlp timed out after ${timeoutMs}ms`));
    }, timeoutMs);
    proc.stdout.setEncoding('utf8');
    proc.stderr.setEncoding('utf8');
    proc.stdout.on('data', (chunk) => {
      stdout = appendChunk(stdout, chunk, 'stdout');
    });
    proc.stderr.on('data', (chunk) => {
      stderr = appendChunk(stderr, chunk, 'stderr');
    });
    proc.once('error', (error) => {
      rejectOnce(error);
    });
    proc.once('close', (code) => {
      if (settled) {
        return;
      }
      if (code === 0) {
        resolveOnce({ stdout, stderr });
        return;
      }
      rejectOnce(new Error(stderr.trim() || `yt-dlp exited with status ${code ?? 'unknown'}`));
    });
  });
}

export function buildYoutubePlaybackResolveArgs(targetUrl: string, format: string): string[] {
  return [YTDLP_SINGLE_VIDEO_ARG, '--get-url', '--no-warnings', '-f', format, targetUrl];
}

export async function resolveYoutubePlaybackUrl(
  targetUrl: string,
  format = DEFAULT_PLAYBACK_FORMAT,
): Promise<string> {
  const { stdout } = await runCapture(
    getYoutubeYtDlpCommand(),
    buildYoutubePlaybackResolveArgs(targetUrl, format),
  );
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

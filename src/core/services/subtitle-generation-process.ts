import { spawn } from 'node:child_process';
import { expandSubtitleGenerationPath } from './subtitle-generation-files';

const OUTPUT_LIMIT = 64 * 1024;

// Keep partial lines between chunks: ffmpeg and whisper both report progress on stderr.
export function runSubtitleGenerationProcess(input: {
  command: string;
  args: string[];
  signal?: AbortSignal;
  onLine?: (line: string) => void;
}): Promise<string> {
  input.signal?.throwIfAborted();
  return new Promise((resolve, reject) => {
    const child = spawn(expandSubtitleGenerationPath(input.command), input.args, {
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    let stdout = '';
    let stderr = '';
    let killTimer: ReturnType<typeof setTimeout> | undefined;
    const abort = () => {
      child.kill('SIGTERM');
      killTimer = setTimeout(() => child.kill('SIGKILL'), 2000);
      killTimer.unref();
    };
    input.signal?.addEventListener('abort', abort, { once: true });
    if (input.signal?.aborted) abort();
    const cleanup = () => {
      input.signal?.removeEventListener('abort', abort);
      clearTimeout(killTimer);
    };
    for (const [stream, isStdout] of [
      [child.stdout, true],
      [child.stderr, false],
    ] as const) {
      let pending = '';
      stream.setEncoding('utf8');
      stream.on('data', (chunk: string) => {
        if (isStdout) stdout = (stdout + chunk).slice(-OUTPUT_LIMIT);
        else stderr = (stderr + chunk).slice(-OUTPUT_LIMIT);
        const lines = (pending + chunk).split(/[\r\n]/);
        pending = (lines.pop() ?? '').slice(-OUTPUT_LIMIT);
        for (const line of lines) input.onLine?.(line);
      });
      stream.on('end', () => {
        if (pending) input.onLine?.(pending);
      });
    }
    child.once('error', (error) => {
      cleanup();
      reject(new Error(`Could not run ${input.command}: ${error.message}`));
    });
    child.once('close', (code) => {
      cleanup();
      if (input.signal?.aborted) reject(new Error('Subtitle generation cancelled.'));
      else if (code !== 0) {
        reject(new Error(`${input.command} exited with status ${code}: ${stderr.trim()}`));
      } else resolve(stdout);
    });
  });
}

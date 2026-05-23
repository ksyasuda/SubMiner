export class ExitSignal extends Error {
  code: number;
  stderr: string;

  constructor(code: number, stderr: string) {
    super(`exit:${code}`);
    this.code = code;
    this.stderr = stderr;
  }
}

function stderrChunkToString(chunk: string | Uint8Array, encoding?: BufferEncoding): string {
  if (typeof chunk === 'string') return chunk;
  return Buffer.from(chunk).toString(encoding);
}

export function withProcessExitIntercept(callback: () => void): ExitSignal {
  const originalExit = process.exit;
  const originalStderrWrite = process.stderr.write;
  let stderr = '';

  try {
    process.stderr.write = ((chunk: string | Uint8Array, ...args: unknown[]) => {
      const encoding = typeof args[0] === 'string' ? (args[0] as BufferEncoding) : undefined;
      stderr += stderrChunkToString(chunk, encoding);
      const writeCallback = args.find((arg): arg is (error?: Error | null) => void => {
        return typeof arg === 'function';
      });
      writeCallback?.();
      return true;
    }) as typeof process.stderr.write;
    process.exit = ((code?: number) => {
      throw new ExitSignal(code ?? 0, stderr);
    }) as typeof process.exit;
    callback();
  } catch (error) {
    if (error instanceof ExitSignal) {
      return error;
    }
    throw error;
  } finally {
    process.stderr.write = originalStderrWrite;
    process.exit = originalExit;
  }

  throw new Error('expected process.exit');
}

type PhaseLogLevel = 'info' | 'warn';

export interface RunLoggedYoutubePhaseOptions {
  startMessage: string;
  finishMessage: string;
  failureMessage?: string;
  log: (level: PhaseLogLevel, message: string) => void;
  now?: () => number;
}

function formatElapsedMs(elapsedMs: number): string {
  const seconds = Math.max(0, elapsedMs) / 1000;
  return `${seconds.toFixed(1)}s`;
}

export async function runLoggedYoutubePhase<T>(
  options: RunLoggedYoutubePhaseOptions,
  run: () => Promise<T>,
): Promise<T> {
  const now = options.now ?? Date.now;
  const startedAt = now();
  options.log('info', options.startMessage);
  try {
    const result = await run();
    options.log('info', `${options.finishMessage} (${formatElapsedMs(now() - startedAt)})`);
    return result;
  } catch (error) {
    const prefix = options.failureMessage ?? options.finishMessage;
    const message = error instanceof Error ? error.message : String(error);
    options.log('warn', `${prefix} after ${formatElapsedMs(now() - startedAt)}: ${message}`);
    throw error;
  }
}

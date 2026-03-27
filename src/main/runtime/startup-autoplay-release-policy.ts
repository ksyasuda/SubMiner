const DEFAULT_AUTOPLAY_RELEASE_RETRY_DELAY_MS = 200;
const STARTUP_AUTOPLAY_RELEASE_TIMEOUT_MS = 15_000;

export function resolveAutoplayReadyMaxReleaseAttempts(options?: {
  forceWhilePaused?: boolean;
  retryDelayMs?: number;
  startupTimeoutMs?: number;
}): number {
  if (options?.forceWhilePaused !== true) {
    return 3;
  }

  const retryDelayMs = Math.max(
    1,
    Math.floor(options.retryDelayMs ?? DEFAULT_AUTOPLAY_RELEASE_RETRY_DELAY_MS),
  );
  const startupTimeoutMs = Math.max(
    retryDelayMs,
    Math.floor(options.startupTimeoutMs ?? STARTUP_AUTOPLAY_RELEASE_TIMEOUT_MS),
  );

  return Math.max(3, Math.ceil(startupTimeoutMs / retryDelayMs));
}

export { DEFAULT_AUTOPLAY_RELEASE_RETRY_DELAY_MS, STARTUP_AUTOPLAY_RELEASE_TIMEOUT_MS };

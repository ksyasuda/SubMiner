const DEFAULT_MAX_PER_MINUTE = 20;
const WINDOW_MS = 60_000;
const SAFETY_REMAINING_THRESHOLD = 5;

export interface AnilistRateLimiter {
  acquire(): Promise<void>;
  recordResponse(headers: Headers): void;
}

export function createAnilistRateLimiter(
  maxPerMinute = DEFAULT_MAX_PER_MINUTE,
): AnilistRateLimiter {
  const timestamps: number[] = [];
  let pauseUntilMs = 0;

  function pruneOld(now: number): void {
    const cutoff = now - WINDOW_MS;
    while (timestamps.length > 0 && timestamps[0]! < cutoff) {
      timestamps.shift();
    }
  }

  return {
    async acquire(): Promise<void> {
      const now = Date.now();

      if (now < pauseUntilMs) {
        const waitMs = pauseUntilMs - now;
        await new Promise((resolve) => setTimeout(resolve, waitMs));
      }

      pruneOld(Date.now());

      if (timestamps.length >= maxPerMinute) {
        const oldest = timestamps[0]!;
        const waitMs = oldest + WINDOW_MS - Date.now() + 100;
        if (waitMs > 0) {
          await new Promise((resolve) => setTimeout(resolve, waitMs));
        }
        pruneOld(Date.now());
      }

      timestamps.push(Date.now());
    },

    recordResponse(headers: Headers): void {
      const remaining = headers.get('x-ratelimit-remaining');
      if (remaining !== null) {
        const n = parseInt(remaining, 10);
        if (Number.isFinite(n) && n < SAFETY_REMAINING_THRESHOLD) {
          const reset = headers.get('x-ratelimit-reset');
          if (reset) {
            const resetMs = parseInt(reset, 10) * 1000;
            if (Number.isFinite(resetMs)) {
              pauseUntilMs = Math.max(pauseUntilMs, resetMs);
            }
          } else {
            pauseUntilMs = Math.max(pauseUntilMs, Date.now() + WINDOW_MS);
          }
        }
      }

      const retryAfter = headers.get('retry-after');
      if (retryAfter) {
        const seconds = parseInt(retryAfter, 10);
        if (Number.isFinite(seconds) && seconds > 0) {
          pauseUntilMs = Math.max(pauseUntilMs, Date.now() + seconds * 1000);
        }
      }
    },
  };
}

/**
 * Waits until mpv has a video window the overlay can attach to.
 *
 * mpv reports a new `path` as soon as loading starts, which for a network
 * stream is seconds before any window exists (idle mpv shows none). A modal
 * opened on the path change lands on a blank desktop with fallback geometry.
 * Ready means the video output is configured (`vo-configured`, a window with a
 * frame in it) and, when a window tracker is running, that it has located the
 * window so overlay geometry follows it.
 */
export interface WaitForPlaybackWindowDeps {
  /** Tracker state, or null when no window tracker is running. */
  isWindowTracked: () => boolean | null;
  /** One-shot mpv property read; may reject while the file is still loading. */
  readProperty: (name: string) => Promise<unknown>;
  wait: (ms: number) => Promise<void>;
  now?: () => number;
  timeoutMs?: number;
  probeIntervalMs?: number;
}

export const DEFAULT_PLAYBACK_WINDOW_TIMEOUT_MS = 20_000;
const DEFAULT_PROBE_INTERVAL_MS = 200;

/** Resolves true once the window is ready, false when the wall-clock budget runs out. */
export async function waitForPlaybackWindow(deps: WaitForPlaybackWindowDeps): Promise<boolean> {
  const now = deps.now ?? Date.now;
  const timeoutMs = deps.timeoutMs ?? DEFAULT_PLAYBACK_WINDOW_TIMEOUT_MS;
  const probeIntervalMs = deps.probeIntervalMs ?? DEFAULT_PROBE_INTERVAL_MS;
  const deadline = now() + timeoutMs;

  for (;;) {
    let voConfigured = false;
    try {
      voConfigured = (await deps.readProperty('vo-configured')) === true;
    } catch {
      // Unreadable between files or before mpv connects; keep polling.
    }
    if (voConfigured && deps.isWindowTracked() !== false) return true;

    const remaining = deadline - now();
    if (remaining <= 0) return false;
    await deps.wait(Math.min(probeIntervalMs, remaining));
  }
}

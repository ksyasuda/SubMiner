/**
 * Confirms that a `loadfile` handed to mpv actually turned into playback.
 *
 * Sending the command proves nothing: mpv accepts the file, fails to decode it
 * (a dead host, a disguised stream the proxy could not fix), fires `end-file`
 * with reason "error", and drops back to `--idle` — with no window, because
 * idle mpv shows none. The UI would happily say "Playing" over a blank desktop.
 *
 * Success is mpv configuring a video output (`vo-configured`), which is
 * literally "a window with frames in it". `file-loaded` is not enough — the
 * broken stream in the wild reached it before dying. Failure is an `end-file`
 * with reason "error"; the end-file of the file being *replaced* arrives with
 * "stop"/"redirect" and is ignored.
 */

export interface PlaybackEndFileEvent {
  reason: string;
  fileError: string | null;
}

export type PlaybackOutcome = { ok: true } | { ok: false; error: string };

export interface WatchPlaybackOutcomeDeps {
  /** Subscribe to mpv end-file events; returns the unsubscribe. */
  onEndFile: (listener: (event: PlaybackEndFileEvent) => void) => () => void;
  /** One-shot mpv property read; may reject while the file is still loading. */
  readProperty: (name: string) => Promise<unknown>;
  wait: (ms: number) => Promise<void>;
  /** Injectable clock; the timeout is wall-clock, not a probe count. */
  now?: () => number;
  timeoutMs?: number;
  probeIntervalMs?: number;
}

export interface PlaybackOutcomeWatch {
  wait: () => Promise<PlaybackOutcome>;
  dispose: () => void;
}

export const DEFAULT_PLAYBACK_OUTCOME_TIMEOUT_MS = 20_000;
const DEFAULT_PROBE_INTERVAL_MS = 500;

/**
 * Call *before* sending `loadfile` so the error subscription cannot lose a
 * race against a fast failure; await `wait()` after the commands went out.
 */
export function watchPlaybackOutcome(deps: WatchPlaybackOutcomeDeps): PlaybackOutcomeWatch {
  const timeoutMs = deps.timeoutMs ?? DEFAULT_PLAYBACK_OUTCOME_TIMEOUT_MS;
  const probeIntervalMs = deps.probeIntervalMs ?? DEFAULT_PROBE_INTERVAL_MS;

  let failure: PlaybackOutcome | null = null;
  const unsubscribe = deps.onEndFile((event) => {
    if (event.reason !== 'error') return;
    failure = {
      ok: false,
      error: event.fileError
        ? `mpv could not play this stream: ${event.fileError}`
        : 'mpv could not play this stream.',
    };
  });

  async function wait(): Promise<PlaybackOutcome> {
    // Wall-clock, not a probe count: a slow `readProperty` must eat into the
    // budget rather than stretch it, and a zero probe interval must still end.
    const now = deps.now ?? Date.now;
    const deadline = now() + timeoutMs;
    while (now() < deadline) {
      if (failure) return failure;
      try {
        if ((await deps.readProperty('vo-configured')) === true) return { ok: true };
      } catch {
        // The property is unreadable while mpv is between files; keep polling.
      }
      if (failure) return failure;
      // Sleeping past the deadline would only delay the timeout report.
      const remaining = deadline - now();
      if (remaining <= 0) break;
      await deps.wait(Math.min(probeIntervalMs, remaining));
    }
    return (
      failure ?? {
        ok: false,
        error: 'Playback did not start. mpv gave no error; try another server or quality.',
      }
    );
  }

  return { wait, dispose: unsubscribe };
}

import { SubsyncManualRunRequest, SubsyncResult } from '../../types';
import { runAutoSubsync } from './subsync-auto';
import { openSubsyncManualPicker, runSubsyncManualUnguarded } from './subsync-manual';
import {
  PLAYBACK_CHANGED_RESULT_MESSAGE,
  STALE_AUTO_RESULT_MESSAGE,
  type SubsyncCoreDeps,
  type TriggerSubsyncFromConfigDeps,
} from './subsync-shared';

// Subsync entry points. Modules: subsync-context (mpv state + session checks), subsync-files
// (extract/write/reload), subsync-candidates (reference discovery + audio), subsync-manual,
// subsync-auto, and subsync-shared (deps, thresholds, formatting).
export { openSubsyncManualPicker } from './subsync-manual';
export type { SubsyncVadPaths, TriggerSubsyncFromConfigDeps } from './subsync-shared';

/** Manual (modal) run, cancelled when playback moves to other media. */
export async function runSubsyncManual(
  request: SubsyncManualRunRequest,
  deps: SubsyncCoreDeps,
): Promise<SubsyncResult> {
  try {
    return await runWithMediaChangeAbort(deps, async (runDeps) => {
      const result = await runSubsyncManualUnguarded(request, runDeps);
      // Tools killed by the abort report generic failures; say what actually happened.
      return !result.ok && runDeps.signal?.aborted
        ? { ok: false, message: PLAYBACK_CHANGED_RESULT_MESSAGE }
        : result;
    });
  } catch (error) {
    if (error instanceof SubsyncCancelledError) {
      return { ok: false, message: PLAYBACK_CHANGED_RESULT_MESSAGE };
    }
    throw error;
  }
}

const MEDIA_CHANGE_POLL_MS = 500;

/** Thrown when a run was cancelled because playback moved to other media. */
export class SubsyncCancelledError extends Error {
  constructor() {
    super(PLAYBACK_CHANGED_RESULT_MESSAGE);
    this.name = 'SubsyncCancelledError';
  }
}

/**
 * Run a subsync task with `deps.signal` aborting once mpv leaves the media the run started on
 * (polls `path`; mpv disconnecting also aborts). Aborting kills spawned ffmpeg/VAD/alass and
 * stops at the next stage boundary, so switching episodes frees the subsync lock right away.
 * A failure caused by the abort surfaces as `SubsyncCancelledError`.
 */
export async function runWithMediaChangeAbort<D extends SubsyncCoreDeps, T>(
  deps: D,
  task: (deps: D) => Promise<T>,
): Promise<T> {
  const client = deps.getMpvClient();
  const initialPath = client?.connected
    ? await client.requestProperty('path').catch(() => null)
    : null;
  const controller = new AbortController();
  let polling = false;
  const timer = setInterval(() => {
    if (polling || controller.signal.aborted) return;
    const current = deps.getMpvClient();
    if (!current?.connected) {
      controller.abort();
      return;
    }
    polling = true;
    current
      .requestProperty('path')
      .then(
        (currentPath) => {
          if (currentPath !== initialPath) controller.abort();
        },
        () => {},
      )
      .finally(() => {
        polling = false;
      });
  }, MEDIA_CHANGE_POLL_MS);
  const signal = deps.signal
    ? AbortSignal.any([deps.signal, controller.signal])
    : controller.signal;
  try {
    return await task({ ...deps, signal });
  } catch (error) {
    if (signal.aborted) throw new SubsyncCancelledError();
    throw error;
  } finally {
    clearInterval(timer);
  }
}

export async function triggerSubsyncFromConfig(deps: TriggerSubsyncFromConfigDeps): Promise<void> {
  if (deps.isSubsyncInProgress()) {
    deps.showMpvOsd('Subsync already running');
    return;
  }

  deps.setSubsyncInProgress(true);
  try {
    await runWithMediaChangeAbort(deps, async (runDeps) => {
      if (runDeps.getResolvedConfig().mode === 'manual') {
        const opened = await openSubsyncManualPicker(runDeps, { skipAudioAnalysis: true });
        if (opened) {
          runDeps.showMpvOsd('Subsync: choose engine and subtitles');
        }
        return;
      }
      await runDeps.runWithSubsyncSpinner(() => runAutoSubsync(runDeps));
    });
  } catch (error) {
    deps.showMpvOsd(
      error instanceof SubsyncCancelledError
        ? STALE_AUTO_RESULT_MESSAGE
        : `Subsync failed: ${(error as Error).message}`,
    );
  } finally {
    deps.setSubsyncInProgress(false);
  }
}

import type { SubtitleData } from '../../types';
import { resolveAutoplayReadyMaxReleaseAttempts } from './startup-autoplay-release-policy';

const PENDING_AUTOPLAY_READY_RETRY_DELAY_MS = 200;
const MAX_PENDING_AUTOPLAY_READY_RETRY_ATTEMPTS = 75;

type MpvClientLike = {
  connected?: boolean;
  requestProperty: (property: string) => Promise<unknown>;
  send: (payload: { command: Array<string | boolean> }) => void;
};

type AutoplayReadyOptions = { forceWhilePaused?: boolean };

export type AutoplayReadySignal = {
  mediaPath: string;
  payload: SubtitleData;
  requestedAtMs: number;
  options?: AutoplayReadyOptions;
};

export type AutoplayReadyGateDeps = {
  isAppOwnedFlowInFlight: () => boolean;
  getCurrentMediaPath: () => string | null;
  getCurrentVideoPath: () => string | null;
  getPlaybackPaused: () => boolean | null;
  getMpvClient: () => MpvClientLike | null;
  signalPluginAutoplayReady: () => void;
  requestOverlayPointerRecovery?: () => void;
  isSignalTargetReady?: (signal: AutoplayReadySignal) => boolean;
  now?: () => number;
  schedule: (callback: () => void, delayMs: number) => ReturnType<typeof setTimeout>;
  logDebug: (message: string) => void;
};

export function createAutoplayReadyGate(deps: AutoplayReadyGateDeps) {
  let autoPlayReadySignalMediaPath: string | null = null;
  let autoPlayReadySignalGeneration = 0;
  let pendingAutoplayReadySignal: AutoplayReadySignal | null = null;
  let pendingAutoplayReadyRetryToken = 0;
  let pendingAutoplayReadyRetryAttempts = 0;
  let scheduledPendingAutoplayReadyRetryToken: number | null = null;
  const now = deps.now ?? (() => Date.now());

  const invalidatePendingAutoplayReadyRetry = (): void => {
    pendingAutoplayReadyRetryToken += 1;
    pendingAutoplayReadyRetryAttempts = 0;
    scheduledPendingAutoplayReadyRetryToken = null;
  };

  const invalidatePendingAutoplayReadyFallbacks = (): void => {
    autoPlayReadySignalMediaPath = null;
    pendingAutoplayReadySignal = null;
    autoPlayReadySignalGeneration += 1;
    invalidatePendingAutoplayReadyRetry();
  };

  const isSignalTargetReady = (signal: AutoplayReadySignal): boolean =>
    deps.isSignalTargetReady?.(signal) ?? true;

  const getSignalMediaPath = (): string =>
    deps.getCurrentMediaPath()?.trim() || deps.getCurrentVideoPath()?.trim() || '__unknown__';

  const markCurrentMediaAutoplayReady = (): void => {
    pendingAutoplayReadySignal = null;
    autoPlayReadySignalMediaPath = getSignalMediaPath();
    autoPlayReadySignalGeneration += 1;
    invalidatePendingAutoplayReadyRetry();
  };

  const setPendingAutoplayReadySignal = (signal: AutoplayReadySignal): boolean => {
    if (
      pendingAutoplayReadySignal &&
      pendingAutoplayReadySignal.mediaPath === signal.mediaPath &&
      pendingAutoplayReadySignal.payload.text === signal.payload.text &&
      pendingAutoplayReadySignal.requestedAtMs <= signal.requestedAtMs
    ) {
      return false;
    }
    pendingAutoplayReadySignal = signal;
    pendingAutoplayReadyRetryAttempts = 0;
    return true;
  };

  const schedulePendingAutoplayReadyRetry = (): void => {
    if (scheduledPendingAutoplayReadyRetryToken === pendingAutoplayReadyRetryToken) {
      return;
    }
    if (pendingAutoplayReadyRetryAttempts >= MAX_PENDING_AUTOPLAY_READY_RETRY_ATTEMPTS) {
      return;
    }

    const retryToken = pendingAutoplayReadyRetryToken;
    pendingAutoplayReadyRetryAttempts += 1;
    scheduledPendingAutoplayReadyRetryToken = retryToken;
    deps.schedule(() => {
      if (scheduledPendingAutoplayReadyRetryToken === retryToken) {
        scheduledPendingAutoplayReadyRetryToken = null;
      }
      if (retryToken !== pendingAutoplayReadyRetryToken || !pendingAutoplayReadySignal) {
        return;
      }
      flushPendingAutoplayReadySignal();
    }, PENDING_AUTOPLAY_READY_RETRY_DELAY_MS);
  };

  const releaseAutoplayReadySignal = (signal: AutoplayReadySignal): void => {
    const mediaPath = signal.mediaPath;
    const releaseRetryDelayMs = 200;
    const maxReleaseAttempts = resolveAutoplayReadyMaxReleaseAttempts({
      forceWhilePaused: signal.options?.forceWhilePaused === true,
      retryDelayMs: releaseRetryDelayMs,
    });
    let releaseUnpauseSent = false;

    const isPlaybackPaused = async (client: MpvClientLike): Promise<boolean> => {
      try {
        const pauseProperty = await client.requestProperty('pause');
        if (typeof pauseProperty === 'boolean') {
          return pauseProperty;
        }
        if (typeof pauseProperty === 'string') {
          return pauseProperty.toLowerCase() !== 'no' && pauseProperty !== '0';
        }
        if (typeof pauseProperty === 'number') {
          return pauseProperty !== 0;
        }
      } catch (error) {
        deps.logDebug(
          `[autoplay-ready] failed to read pause property for media ${mediaPath}: ${
            error instanceof Error ? error.message : String(error)
          }`,
        );
      }

      return true;
    };

    const attemptRelease = (playbackGeneration: number, attempt: number): void => {
      void (async () => {
        if (
          autoPlayReadySignalMediaPath !== mediaPath ||
          playbackGeneration !== autoPlayReadySignalGeneration
        ) {
          return;
        }

        const mpvClient = deps.getMpvClient();
        if (!mpvClient?.connected) {
          if (attempt < maxReleaseAttempts) {
            deps.schedule(
              () => attemptRelease(playbackGeneration, attempt + 1),
              releaseRetryDelayMs,
            );
          }
          return;
        }

        if (releaseUnpauseSent && deps.getPlaybackPaused() === true) {
          deps.logDebug(
            `[autoplay-ready] stopped release retries after playback paused again for media ${mediaPath}`,
          );
          return;
        }

        const shouldUnpause = await isPlaybackPaused(mpvClient);
        if (!shouldUnpause) {
          return;
        }

        mpvClient.send({ command: ['set_property', 'pause', false] });
        releaseUnpauseSent = true;
        if (attempt < maxReleaseAttempts) {
          deps.schedule(() => attemptRelease(playbackGeneration, attempt + 1), releaseRetryDelayMs);
        }
      })();
    };

    pendingAutoplayReadySignal = null;
    invalidatePendingAutoplayReadyRetry();
    autoPlayReadySignalMediaPath = mediaPath;
    const playbackGeneration = ++autoPlayReadySignalGeneration;
    deps.signalPluginAutoplayReady();
    deps.requestOverlayPointerRecovery?.();
    attemptRelease(playbackGeneration, 0);
  };

  const maybeReleaseAutoplayReadySignal = (signal: AutoplayReadySignal): void => {
    if (autoPlayReadySignalMediaPath === signal.mediaPath) {
      pendingAutoplayReadySignal = null;
      return;
    }
    if (!isSignalTargetReady(signal)) {
      const pendingSignalChanged = setPendingAutoplayReadySignal(signal);
      schedulePendingAutoplayReadyRetry();
      if (pendingSignalChanged) {
        deps.logDebug(
          `[autoplay-ready] deferred until signal target is ready for media ${signal.mediaPath}`,
        );
      }
      return;
    }

    releaseAutoplayReadySignal(signal);
  };

  const maybeSignalPluginAutoplayReady = (
    payload: SubtitleData,
    options?: AutoplayReadyOptions,
  ): void => {
    if (deps.isAppOwnedFlowInFlight()) {
      deps.logDebug('[autoplay-ready] suppressed while app-owned YouTube flow is active');
      return;
    }
    if (!payload.text.trim()) {
      return;
    }

    maybeReleaseAutoplayReadySignal({
      mediaPath: getSignalMediaPath(),
      payload,
      requestedAtMs: now(),
      options,
    });
  };

  const flushPendingAutoplayReadySignal = (): void => {
    if (!pendingAutoplayReadySignal) {
      return;
    }

    const pendingSignal = pendingAutoplayReadySignal;
    if (getSignalMediaPath() !== pendingSignal.mediaPath) {
      pendingAutoplayReadySignal = null;
      deps.logDebug(
        `[autoplay-ready] dropped deferred signal for stale media ${pendingSignal.mediaPath}`,
      );
      return;
    }
    maybeReleaseAutoplayReadySignal(pendingSignal);
  };

  return {
    flushPendingAutoplayReadySignal,
    getAutoPlayReadySignalMediaPath: (): string | null => autoPlayReadySignalMediaPath,
    invalidatePendingAutoplayReadyFallbacks,
    markCurrentMediaAutoplayReady,
    maybeSignalPluginAutoplayReady,
  };
}

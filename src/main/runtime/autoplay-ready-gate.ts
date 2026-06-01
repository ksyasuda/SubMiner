import type { SubtitleData } from '../../types';
import { resolveAutoplayReadyMaxReleaseAttempts } from './startup-autoplay-release-policy';

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
  const now = deps.now ?? (() => Date.now());

  const invalidatePendingAutoplayReadyFallbacks = (): void => {
    autoPlayReadySignalMediaPath = null;
    pendingAutoplayReadySignal = null;
    autoPlayReadySignalGeneration += 1;
  };

  const isSignalTargetReady = (signal: AutoplayReadySignal): boolean =>
    deps.isSignalTargetReady?.(signal) ?? true;

  const getSignalMediaPath = (): string =>
    deps.getCurrentMediaPath()?.trim() || deps.getCurrentVideoPath()?.trim() || '__unknown__';

  const markCurrentMediaAutoplayReady = (): void => {
    pendingAutoplayReadySignal = null;
    autoPlayReadySignalMediaPath = getSignalMediaPath();
    autoPlayReadySignalGeneration += 1;
  };

  const setPendingAutoplayReadySignal = (signal: AutoplayReadySignal): void => {
    if (
      pendingAutoplayReadySignal &&
      pendingAutoplayReadySignal.mediaPath === signal.mediaPath &&
      pendingAutoplayReadySignal.payload.text === signal.payload.text &&
      pendingAutoplayReadySignal.requestedAtMs <= signal.requestedAtMs
    ) {
      return;
    }
    pendingAutoplayReadySignal = signal;
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
      setPendingAutoplayReadySignal(signal);
      deps.logDebug(
        `[autoplay-ready] deferred until signal target is ready for media ${signal.mediaPath}`,
      );
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

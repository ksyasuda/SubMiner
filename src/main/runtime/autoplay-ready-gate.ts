import type { SubtitleData } from '../../types';
import { resolveAutoplayReadyMaxReleaseAttempts } from './startup-autoplay-release-policy';

type MpvClientLike = {
  connected?: boolean;
  requestProperty: (property: string) => Promise<unknown>;
  send: (payload: { command: Array<string | boolean> }) => void;
};

export type AutoplayReadyGateDeps = {
  isAppOwnedFlowInFlight: () => boolean;
  getCurrentMediaPath: () => string | null;
  getCurrentVideoPath: () => string | null;
  getPlaybackPaused: () => boolean | null;
  getMpvClient: () => MpvClientLike | null;
  signalPluginAutoplayReady: () => void;
  isSignalTargetReady?: () => boolean;
  schedule: (callback: () => void, delayMs: number) => ReturnType<typeof setTimeout>;
  logDebug: (message: string) => void;
};

export function createAutoplayReadyGate(deps: AutoplayReadyGateDeps) {
  let autoPlayReadySignalMediaPath: string | null = null;
  let autoPlayReadySignalGeneration = 0;
  let pendingAutoplayReadySignal: {
    payload: SubtitleData;
    options?: { forceWhilePaused?: boolean };
  } | null = null;

  const invalidatePendingAutoplayReadyFallbacks = (): void => {
    autoPlayReadySignalMediaPath = null;
    pendingAutoplayReadySignal = null;
    autoPlayReadySignalGeneration += 1;
  };

  const isSignalTargetReady = (): boolean => deps.isSignalTargetReady?.() ?? true;

  const maybeSignalPluginAutoplayReady = (
    payload: SubtitleData,
    options?: { forceWhilePaused?: boolean },
  ): void => {
    if (deps.isAppOwnedFlowInFlight()) {
      deps.logDebug('[autoplay-ready] suppressed while app-owned YouTube flow is active');
      return;
    }
    if (!payload.text.trim()) {
      return;
    }

    const mediaPath =
      deps.getCurrentMediaPath()?.trim() || deps.getCurrentVideoPath()?.trim() || '__unknown__';
    const duplicateMediaSignal = autoPlayReadySignalMediaPath === mediaPath;
    const releaseRetryDelayMs = 200;
    const maxReleaseAttempts = resolveAutoplayReadyMaxReleaseAttempts({
      forceWhilePaused: options?.forceWhilePaused === true,
      retryDelayMs: releaseRetryDelayMs,
    });

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

        const shouldUnpause = await isPlaybackPaused(mpvClient);
        if (!shouldUnpause) {
          return;
        }

        mpvClient.send({ command: ['set_property', 'pause', false] });
        if (attempt < maxReleaseAttempts) {
          deps.schedule(() => attemptRelease(playbackGeneration, attempt + 1), releaseRetryDelayMs);
        }
      })();
    };

    if (duplicateMediaSignal) {
      pendingAutoplayReadySignal = null;
      return;
    }
    if (!isSignalTargetReady()) {
      pendingAutoplayReadySignal = { payload, options };
      deps.logDebug(
        `[autoplay-ready] deferred until signal target is ready for media ${mediaPath}`,
      );
      return;
    }

    pendingAutoplayReadySignal = null;
    autoPlayReadySignalMediaPath = mediaPath;
    const playbackGeneration = ++autoPlayReadySignalGeneration;
    deps.signalPluginAutoplayReady();
    attemptRelease(playbackGeneration, 0);
  };

  const flushPendingAutoplayReadySignal = (): void => {
    if (!pendingAutoplayReadySignal || !isSignalTargetReady()) {
      return;
    }

    const pendingSignal = pendingAutoplayReadySignal;
    pendingAutoplayReadySignal = null;
    maybeSignalPluginAutoplayReady(pendingSignal.payload, pendingSignal.options);
  };

  return {
    flushPendingAutoplayReadySignal,
    getAutoPlayReadySignalMediaPath: (): string | null => autoPlayReadySignalMediaPath,
    invalidatePendingAutoplayReadyFallbacks,
    maybeSignalPluginAutoplayReady,
  };
}

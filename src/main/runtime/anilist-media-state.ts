import type { AnilistMediaGuessRuntimeState } from './anilist-media-guess';

export function createGetCurrentAnilistMediaKeyHandler(deps: {
  getCurrentMediaPath: () => string | null;
}) {
  return (): string | null => {
    const mediaPath = deps.getCurrentMediaPath()?.trim();
    return mediaPath && mediaPath.length > 0 ? mediaPath : null;
  };
}

export function createResetAnilistMediaTrackingHandler(deps: {
  setMediaKey: (value: string | null) => void;
  setMediaDurationSec: (value: number | null) => void;
  setMediaGuess: (value: AnilistMediaGuessRuntimeState['mediaGuess']) => void;
  setMediaGuessPromise: (value: AnilistMediaGuessRuntimeState['mediaGuessPromise']) => void;
  setLastDurationProbeAtMs: (value: number) => void;
}) {
  return (mediaKey: string | null): void => {
    deps.setMediaKey(mediaKey);
    deps.setMediaDurationSec(null);
    deps.setMediaGuess(null);
    deps.setMediaGuessPromise(null);
    deps.setLastDurationProbeAtMs(0);
  };
}

export function createGetAnilistMediaGuessRuntimeStateHandler(deps: {
  getMediaKey: () => string | null;
  getMediaDurationSec: () => number | null;
  getMediaGuess: () => AnilistMediaGuessRuntimeState['mediaGuess'];
  getMediaGuessPromise: () => AnilistMediaGuessRuntimeState['mediaGuessPromise'];
  getLastDurationProbeAtMs: () => number;
}) {
  return (): AnilistMediaGuessRuntimeState => ({
    mediaKey: deps.getMediaKey(),
    mediaDurationSec: deps.getMediaDurationSec(),
    mediaGuess: deps.getMediaGuess(),
    mediaGuessPromise: deps.getMediaGuessPromise(),
    lastDurationProbeAtMs: deps.getLastDurationProbeAtMs(),
  });
}

export function createSetAnilistMediaGuessRuntimeStateHandler(deps: {
  setMediaKey: (value: string | null) => void;
  setMediaDurationSec: (value: number | null) => void;
  setMediaGuess: (value: AnilistMediaGuessRuntimeState['mediaGuess']) => void;
  setMediaGuessPromise: (value: AnilistMediaGuessRuntimeState['mediaGuessPromise']) => void;
  setLastDurationProbeAtMs: (value: number) => void;
}) {
  return (state: AnilistMediaGuessRuntimeState): void => {
    deps.setMediaKey(state.mediaKey);
    deps.setMediaDurationSec(state.mediaDurationSec);
    deps.setMediaGuess(state.mediaGuess);
    deps.setMediaGuessPromise(state.mediaGuessPromise);
    deps.setLastDurationProbeAtMs(state.lastDurationProbeAtMs);
  };
}

export function createResetAnilistMediaGuessStateHandler(deps: {
  setMediaGuess: (value: AnilistMediaGuessRuntimeState['mediaGuess']) => void;
  setMediaGuessPromise: (value: AnilistMediaGuessRuntimeState['mediaGuessPromise']) => void;
}) {
  return (): void => {
    deps.setMediaGuess(null);
    deps.setMediaGuessPromise(null);
  };
}

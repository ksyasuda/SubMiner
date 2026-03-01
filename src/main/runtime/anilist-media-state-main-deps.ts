import type {
  createGetAnilistMediaGuessRuntimeStateHandler,
  createGetCurrentAnilistMediaKeyHandler,
  createResetAnilistMediaGuessStateHandler,
  createResetAnilistMediaTrackingHandler,
  createSetAnilistMediaGuessRuntimeStateHandler,
} from './anilist-media-state';

type GetCurrentAnilistMediaKeyMainDeps = Parameters<
  typeof createGetCurrentAnilistMediaKeyHandler
>[0];
type ResetAnilistMediaTrackingMainDeps = Parameters<
  typeof createResetAnilistMediaTrackingHandler
>[0];
type GetAnilistMediaGuessRuntimeStateMainDeps = Parameters<
  typeof createGetAnilistMediaGuessRuntimeStateHandler
>[0];
type SetAnilistMediaGuessRuntimeStateMainDeps = Parameters<
  typeof createSetAnilistMediaGuessRuntimeStateHandler
>[0];
type ResetAnilistMediaGuessStateMainDeps = Parameters<
  typeof createResetAnilistMediaGuessStateHandler
>[0];

export function createBuildGetCurrentAnilistMediaKeyMainDepsHandler(
  deps: GetCurrentAnilistMediaKeyMainDeps,
) {
  return (): GetCurrentAnilistMediaKeyMainDeps => ({
    getCurrentMediaPath: () => deps.getCurrentMediaPath(),
  });
}

export function createBuildResetAnilistMediaTrackingMainDepsHandler(
  deps: ResetAnilistMediaTrackingMainDeps,
) {
  return (): ResetAnilistMediaTrackingMainDeps => ({
    setMediaKey: (value) => deps.setMediaKey(value),
    setMediaDurationSec: (value) => deps.setMediaDurationSec(value),
    setMediaGuess: (value) => deps.setMediaGuess(value),
    setMediaGuessPromise: (value) => deps.setMediaGuessPromise(value),
    setLastDurationProbeAtMs: (value: number) => deps.setLastDurationProbeAtMs(value),
  });
}

export function createBuildGetAnilistMediaGuessRuntimeStateMainDepsHandler(
  deps: GetAnilistMediaGuessRuntimeStateMainDeps,
) {
  return (): GetAnilistMediaGuessRuntimeStateMainDeps => ({
    getMediaKey: () => deps.getMediaKey(),
    getMediaDurationSec: () => deps.getMediaDurationSec(),
    getMediaGuess: () => deps.getMediaGuess(),
    getMediaGuessPromise: () => deps.getMediaGuessPromise(),
    getLastDurationProbeAtMs: () => deps.getLastDurationProbeAtMs(),
  });
}

export function createBuildSetAnilistMediaGuessRuntimeStateMainDepsHandler(
  deps: SetAnilistMediaGuessRuntimeStateMainDeps,
) {
  return (): SetAnilistMediaGuessRuntimeStateMainDeps => ({
    setMediaKey: (value) => deps.setMediaKey(value),
    setMediaDurationSec: (value) => deps.setMediaDurationSec(value),
    setMediaGuess: (value) => deps.setMediaGuess(value),
    setMediaGuessPromise: (value) => deps.setMediaGuessPromise(value),
    setLastDurationProbeAtMs: (value: number) => deps.setLastDurationProbeAtMs(value),
  });
}

export function createBuildResetAnilistMediaGuessStateMainDepsHandler(
  deps: ResetAnilistMediaGuessStateMainDeps,
) {
  return (): ResetAnilistMediaGuessStateMainDeps => ({
    setMediaGuess: (value) => deps.setMediaGuess(value),
    setMediaGuessPromise: (value) => deps.setMediaGuessPromise(value),
  });
}

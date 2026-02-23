import type {
  createEnsureAnilistMediaGuessHandler,
  createMaybeProbeAnilistDurationHandler,
} from './anilist-media-guess';

type MaybeProbeAnilistDurationMainDeps = Parameters<typeof createMaybeProbeAnilistDurationHandler>[0];
type EnsureAnilistMediaGuessMainDeps = Parameters<typeof createEnsureAnilistMediaGuessHandler>[0];

export function createBuildMaybeProbeAnilistDurationMainDepsHandler(
  deps: MaybeProbeAnilistDurationMainDeps,
) {
  return (): MaybeProbeAnilistDurationMainDeps => ({
    getState: () => deps.getState(),
    setState: (state) => deps.setState(state),
    durationRetryIntervalMs: deps.durationRetryIntervalMs,
    now: () => deps.now(),
    requestMpvDuration: () => deps.requestMpvDuration(),
    logWarn: (message: string, error: unknown) => deps.logWarn(message, error),
  });
}

export function createBuildEnsureAnilistMediaGuessMainDepsHandler(deps: EnsureAnilistMediaGuessMainDeps) {
  return (): EnsureAnilistMediaGuessMainDeps => ({
    getState: () => deps.getState(),
    setState: (state) => deps.setState(state),
    resolveMediaPathForJimaku: (currentMediaPath) => deps.resolveMediaPathForJimaku(currentMediaPath),
    getCurrentMediaPath: () => deps.getCurrentMediaPath(),
    getCurrentMediaTitle: () => deps.getCurrentMediaTitle(),
    guessAnilistMediaInfo: (mediaPath, mediaTitle) => deps.guessAnilistMediaInfo(mediaPath, mediaTitle),
  });
}

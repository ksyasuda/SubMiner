import type {
  createMaybeRunAnilistPostWatchUpdateHandler,
  createProcessNextAnilistRetryUpdateHandler,
} from './anilist-post-watch';

type ProcessNextAnilistRetryUpdateMainDeps = Parameters<
  typeof createProcessNextAnilistRetryUpdateHandler
>[0];
type MaybeRunAnilistPostWatchUpdateMainDeps = Parameters<
  typeof createMaybeRunAnilistPostWatchUpdateHandler
>[0];

export function createBuildProcessNextAnilistRetryUpdateMainDepsHandler(
  deps: ProcessNextAnilistRetryUpdateMainDeps,
) {
  return (): ProcessNextAnilistRetryUpdateMainDeps => ({
    nextReady: () => deps.nextReady(),
    refreshRetryQueueState: () => deps.refreshRetryQueueState(),
    setLastAttemptAt: (value: number) => deps.setLastAttemptAt(value),
    setLastError: (value: string | null) => deps.setLastError(value),
    refreshAnilistClientSecretState: () => deps.refreshAnilistClientSecretState(),
    updateAnilistPostWatchProgress: (accessToken: string, title: string, episode: number) =>
      deps.updateAnilistPostWatchProgress(accessToken, title, episode),
    markSuccess: (key: string) => deps.markSuccess(key),
    rememberAttemptedUpdateKey: (key: string) => deps.rememberAttemptedUpdateKey(key),
    markFailure: (key: string, message: string) => deps.markFailure(key, message),
    logInfo: (message: string) => deps.logInfo(message),
    now: () => deps.now(),
  });
}

export function createBuildMaybeRunAnilistPostWatchUpdateMainDepsHandler(
  deps: MaybeRunAnilistPostWatchUpdateMainDeps,
) {
  return (): MaybeRunAnilistPostWatchUpdateMainDeps => ({
    getInFlight: () => deps.getInFlight(),
    setInFlight: (value: boolean) => deps.setInFlight(value),
    getResolvedConfig: () => deps.getResolvedConfig(),
    isAnilistTrackingEnabled: (config) => deps.isAnilistTrackingEnabled(config),
    getCurrentMediaKey: () => deps.getCurrentMediaKey(),
    hasMpvClient: () => deps.hasMpvClient(),
    getTrackedMediaKey: () => deps.getTrackedMediaKey(),
    resetTrackedMedia: (mediaKey) => deps.resetTrackedMedia(mediaKey),
    getWatchedSeconds: () => deps.getWatchedSeconds(),
    maybeProbeAnilistDuration: (mediaKey: string) => deps.maybeProbeAnilistDuration(mediaKey),
    ensureAnilistMediaGuess: (mediaKey: string) => deps.ensureAnilistMediaGuess(mediaKey),
    hasAttemptedUpdateKey: (key: string) => deps.hasAttemptedUpdateKey(key),
    processNextAnilistRetryUpdate: () => deps.processNextAnilistRetryUpdate(),
    refreshAnilistClientSecretState: () => deps.refreshAnilistClientSecretState(),
    enqueueRetry: (key: string, title: string, episode: number) =>
      deps.enqueueRetry(key, title, episode),
    markRetryFailure: (key: string, message: string) => deps.markRetryFailure(key, message),
    markRetrySuccess: (key: string) => deps.markRetrySuccess(key),
    refreshRetryQueueState: () => deps.refreshRetryQueueState(),
    updateAnilistPostWatchProgress: (accessToken: string, title: string, episode: number) =>
      deps.updateAnilistPostWatchProgress(accessToken, title, episode),
    rememberAttemptedUpdateKey: (key: string) => deps.rememberAttemptedUpdateKey(key),
    showMpvOsd: (message: string) => deps.showMpvOsd(message),
    logInfo: (message: string) => deps.logInfo(message),
    logWarn: (message: string) => deps.logWarn(message),
    minWatchSeconds: deps.minWatchSeconds,
    minWatchRatio: deps.minWatchRatio,
  });
}

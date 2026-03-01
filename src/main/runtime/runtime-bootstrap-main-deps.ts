import type { AnilistStateRuntimeDeps } from './anilist-state';
import type { ConfigDerivedRuntimeDeps } from './config-derived';
import type { ImmersionMediaRuntimeDeps } from './immersion-media';
import type { MainSubsyncRuntimeDeps } from './subsync-runtime';

export function createBuildImmersionMediaRuntimeMainDepsHandler(deps: ImmersionMediaRuntimeDeps) {
  return (): ImmersionMediaRuntimeDeps => ({
    getResolvedConfig: () => deps.getResolvedConfig(),
    defaultImmersionDbPath: deps.defaultImmersionDbPath,
    getTracker: () => deps.getTracker(),
    getMpvClient: () => deps.getMpvClient(),
    getCurrentMediaPath: () => deps.getCurrentMediaPath(),
    getCurrentMediaTitle: () => deps.getCurrentMediaTitle(),
    sleep: deps.sleep,
    seedWaitMs: deps.seedWaitMs,
    seedAttempts: deps.seedAttempts,
    logDebug: (message: string) => deps.logDebug(message),
    logInfo: (message: string) => deps.logInfo(message),
  });
}

export function createBuildAnilistStateRuntimeMainDepsHandler(deps: AnilistStateRuntimeDeps) {
  return (): AnilistStateRuntimeDeps => ({
    getClientSecretState: () => deps.getClientSecretState(),
    setClientSecretState: (next) => deps.setClientSecretState(next),
    getRetryQueueState: () => deps.getRetryQueueState(),
    setRetryQueueState: (next) => deps.setRetryQueueState(next),
    getUpdateQueueSnapshot: () => deps.getUpdateQueueSnapshot(),
    clearStoredToken: () => deps.clearStoredToken(),
    clearCachedAccessToken: () => deps.clearCachedAccessToken(),
  });
}

export function createBuildConfigDerivedRuntimeMainDepsHandler(deps: ConfigDerivedRuntimeDeps) {
  return (): ConfigDerivedRuntimeDeps => ({
    getResolvedConfig: () => deps.getResolvedConfig(),
    getRuntimeOptionsManager: () => deps.getRuntimeOptionsManager(),
    defaultJimakuLanguagePreference: deps.defaultJimakuLanguagePreference,
    defaultJimakuMaxEntryResults: deps.defaultJimakuMaxEntryResults,
    defaultJimakuApiBaseUrl: deps.defaultJimakuApiBaseUrl,
  });
}

export function createBuildMainSubsyncRuntimeMainDepsHandler(deps: MainSubsyncRuntimeDeps) {
  return (): MainSubsyncRuntimeDeps => ({
    getMpvClient: () => deps.getMpvClient(),
    getResolvedConfig: () => deps.getResolvedConfig(),
    getSubsyncInProgress: () => deps.getSubsyncInProgress(),
    setSubsyncInProgress: (inProgress: boolean) => deps.setSubsyncInProgress(inProgress),
    showMpvOsd: (text: string) => deps.showMpvOsd(text),
    openManualPicker: (payload) => deps.openManualPicker(payload),
  });
}

import type {
  createLaunchBackgroundWarmupTaskHandler,
  createStartBackgroundWarmupsHandler,
} from './startup-warmups';

type LaunchBackgroundWarmupTaskMainDeps = Parameters<
  typeof createLaunchBackgroundWarmupTaskHandler
>[0];
type StartBackgroundWarmupsMainDeps = Parameters<typeof createStartBackgroundWarmupsHandler>[0];

export function createBuildLaunchBackgroundWarmupTaskMainDepsHandler(
  deps: LaunchBackgroundWarmupTaskMainDeps,
) {
  return (): LaunchBackgroundWarmupTaskMainDeps => ({
    now: () => deps.now(),
    logDebug: (message: string) => deps.logDebug(message),
    logWarn: (message: string) => deps.logWarn(message),
  });
}

export function createBuildStartBackgroundWarmupsMainDepsHandler(
  deps: StartBackgroundWarmupsMainDeps,
) {
  return (): StartBackgroundWarmupsMainDeps => ({
    getStarted: () => deps.getStarted(),
    setStarted: (started: boolean) => deps.setStarted(started),
    isTexthookerOnlyMode: () => deps.isTexthookerOnlyMode(),
    launchTask: (label: string, task: () => Promise<void>) => deps.launchTask(label, task),
    createMecabTokenizerAndCheck: () => deps.createMecabTokenizerAndCheck(),
    ensureYomitanExtensionLoaded: () => deps.ensureYomitanExtensionLoaded(),
    prewarmSubtitleDictionaries: () => deps.prewarmSubtitleDictionaries(),
    shouldWarmupMecab: () => deps.shouldWarmupMecab(),
    shouldWarmupYomitanExtension: () => deps.shouldWarmupYomitanExtension(),
    shouldWarmupSubtitleDictionaries: () => deps.shouldWarmupSubtitleDictionaries(),
    shouldWarmupJellyfinRemoteSession: () => deps.shouldWarmupJellyfinRemoteSession(),
    shouldAutoConnectJellyfinRemote: () => deps.shouldAutoConnectJellyfinRemote(),
    startJellyfinRemoteSession: () => deps.startJellyfinRemoteSession(),
    logDebug: deps.logDebug,
  });
}

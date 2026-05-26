import type {
  EnsureMpvConnectedDeps,
  LaunchMpvForJellyfinDeps,
  WaitForMpvConnectedDeps,
} from './jellyfin-remote-connection';

export function createBuildWaitForMpvConnectedMainDepsHandler(deps: WaitForMpvConnectedDeps) {
  return (): WaitForMpvConnectedDeps => ({
    getMpvClient: () => deps.getMpvClient(),
    now: () => deps.now(),
    sleep: (delayMs: number) => deps.sleep(delayMs),
  });
}

export function createBuildLaunchMpvIdleForJellyfinPlaybackMainDepsHandler(
  deps: LaunchMpvForJellyfinDeps,
) {
  return (): LaunchMpvForJellyfinDeps => ({
    getSocketPath: () => deps.getSocketPath(),
    getLaunchMode: () => deps.getLaunchMode(),
    platform: deps.platform,
    execPath: deps.execPath,
    getRuntimePluginEntrypoint: deps.getRuntimePluginEntrypoint,
    getInstalledPluginDetection: deps.getInstalledPluginDetection,
    getPluginRuntimeConfig: deps.getPluginRuntimeConfig,
    getDefaultMpvLogPath: () => deps.getDefaultMpvLogPath(),
    defaultMpvArgs: deps.defaultMpvArgs,
    removeSocketPath: (socketPath: string) => deps.removeSocketPath(socketPath),
    spawnMpv: (args: string[]) => deps.spawnMpv(args),
    logWarn: (message: string, error: unknown) => deps.logWarn(message, error),
    logInfo: (message: string) => deps.logInfo(message),
  });
}

export function createBuildEnsureMpvConnectedForJellyfinPlaybackMainDepsHandler(
  deps: EnsureMpvConnectedDeps,
) {
  return (): EnsureMpvConnectedDeps => ({
    getMpvClient: () => deps.getMpvClient(),
    setMpvClient: (client) => deps.setMpvClient(client),
    createMpvClient: () => deps.createMpvClient(),
    waitForMpvConnected: (timeoutMs: number) => deps.waitForMpvConnected(timeoutMs),
    launchMpvIdleForJellyfinPlayback: () => deps.launchMpvIdleForJellyfinPlayback(),
    getAutoLaunchInFlight: () => deps.getAutoLaunchInFlight(),
    setAutoLaunchInFlight: (promise) => deps.setAutoLaunchInFlight(promise),
    connectTimeoutMs: deps.connectTimeoutMs,
    autoLaunchTimeoutMs: deps.autoLaunchTimeoutMs,
  });
}

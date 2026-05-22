import type {
  JellyfinRemoteGeneralCommandHandlerDeps,
  JellyfinRemotePlayHandlerDeps,
  JellyfinRemotePlaystateHandlerDeps,
} from './jellyfin-remote-commands';
import type {
  JellyfinRemoteProgressReporterDeps,
  JellyfinRemoteStoppedReporterDeps,
} from './jellyfin-remote-playback';

export function createBuildHandleJellyfinRemotePlayMainDepsHandler(
  deps: JellyfinRemotePlayHandlerDeps,
) {
  return (): JellyfinRemotePlayHandlerDeps => ({
    getConfiguredSession: () => deps.getConfiguredSession(),
    getClientInfo: () => deps.getClientInfo(),
    getJellyfinConfig: () => deps.getJellyfinConfig(),
    ...(deps.getActivePlayback
      ? { getActivePlayback: () => deps.getActivePlayback?.() ?? null }
      : {}),
    playJellyfinItem: (params) => deps.playJellyfinItem(params),
    logWarn: (message: string) => deps.logWarn(message),
  });
}

export function createBuildHandleJellyfinRemotePlaystateMainDepsHandler(
  deps: JellyfinRemotePlaystateHandlerDeps,
) {
  return (): JellyfinRemotePlaystateHandlerDeps => ({
    getMpvClient: () => deps.getMpvClient(),
    sendMpvCommand: (client, command) => deps.sendMpvCommand(client, command),
    reportJellyfinRemoteProgress: (force: boolean) => deps.reportJellyfinRemoteProgress(force),
    reportJellyfinRemoteStopped: () => deps.reportJellyfinRemoteStopped(),
    jellyfinTicksToSeconds: (ticks: number) => deps.jellyfinTicksToSeconds(ticks),
  });
}

export function createBuildHandleJellyfinRemoteGeneralCommandMainDepsHandler(
  deps: JellyfinRemoteGeneralCommandHandlerDeps,
) {
  return (): JellyfinRemoteGeneralCommandHandlerDeps => ({
    getMpvClient: () => deps.getMpvClient(),
    sendMpvCommand: (client, command) => deps.sendMpvCommand(client, command),
    getActivePlayback: () => deps.getActivePlayback(),
    reportJellyfinRemoteProgress: (force: boolean) => deps.reportJellyfinRemoteProgress(force),
    logDebug: (message: string) => deps.logDebug(message),
  });
}

export function createBuildReportJellyfinRemoteProgressMainDepsHandler(
  deps: JellyfinRemoteProgressReporterDeps,
) {
  return (): JellyfinRemoteProgressReporterDeps => ({
    getActivePlayback: () => deps.getActivePlayback(),
    clearActivePlayback: () => deps.clearActivePlayback(),
    getSession: () => deps.getSession(),
    getMpvClient: () => deps.getMpvClient(),
    getNow: () => deps.getNow(),
    getLastProgressAtMs: () => deps.getLastProgressAtMs(),
    setLastProgressAtMs: (value: number) => deps.setLastProgressAtMs(value),
    progressIntervalMs: deps.progressIntervalMs,
    ticksPerSecond: deps.ticksPerSecond,
    logDebug: (message: string, error: unknown) => deps.logDebug(message, error),
  });
}

export function createBuildReportJellyfinRemoteStoppedMainDepsHandler(
  deps: JellyfinRemoteStoppedReporterDeps,
) {
  return (): JellyfinRemoteStoppedReporterDeps => ({
    getActivePlayback: () => deps.getActivePlayback(),
    clearActivePlayback: () => deps.clearActivePlayback(),
    getSession: () => deps.getSession(),
    logDebug: (message: string, error: unknown) => deps.logDebug(message, error),
  });
}

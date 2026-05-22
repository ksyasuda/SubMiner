import type {
  createStartJellyfinRemoteSessionHandler,
  createStopJellyfinRemoteSessionHandler,
} from './jellyfin-remote-session-lifecycle';

type StartJellyfinRemoteSessionMainDeps = Parameters<
  typeof createStartJellyfinRemoteSessionHandler
>[0];
type StopJellyfinRemoteSessionMainDeps = Parameters<
  typeof createStopJellyfinRemoteSessionHandler
>[0];

export function createBuildStartJellyfinRemoteSessionMainDepsHandler(
  deps: StartJellyfinRemoteSessionMainDeps,
) {
  return (): StartJellyfinRemoteSessionMainDeps => ({
    getJellyfinConfig: () => deps.getJellyfinConfig(),
    getCurrentSession: () => deps.getCurrentSession(),
    setCurrentSession: (session) => deps.setCurrentSession(session),
    createRemoteSessionService: (options) => deps.createRemoteSessionService(options),
    getClientInfo: () => deps.getClientInfo(),
    getHostName: () => deps.getHostName(),
    defaultDeviceId: deps.defaultDeviceId,
    defaultClientName: deps.defaultClientName,
    defaultClientVersion: deps.defaultClientVersion,
    handlePlay: (payload) => deps.handlePlay(payload),
    handlePlaystate: (payload) => deps.handlePlaystate(payload),
    handleGeneralCommand: (payload) => deps.handleGeneralCommand(payload),
    logInfo: (message: string) => deps.logInfo(message),
    logWarn: (message: string, details?: unknown) => deps.logWarn(message, details),
    onSessionStateChanged: deps.onSessionStateChanged,
  });
}

export function createBuildStopJellyfinRemoteSessionMainDepsHandler(
  deps: StopJellyfinRemoteSessionMainDeps,
) {
  return (): StopJellyfinRemoteSessionMainDeps => ({
    getCurrentSession: () => deps.getCurrentSession(),
    setCurrentSession: (session) => deps.setCurrentSession(session),
    clearActivePlayback: () => deps.clearActivePlayback(),
    onSessionStateChanged: deps.onSessionStateChanged,
  });
}

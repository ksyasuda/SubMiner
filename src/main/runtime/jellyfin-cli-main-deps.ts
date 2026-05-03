import type { createHandleJellyfinAuthCommands } from './jellyfin-cli-auth';
import type { createHandleJellyfinListCommands } from './jellyfin-cli-list';
import type { createHandleJellyfinPlayCommand } from './jellyfin-cli-play';
import type { createHandleJellyfinRemoteAnnounceCommand } from './jellyfin-cli-remote-announce';

type HandleJellyfinAuthCommandsMainDeps = Parameters<typeof createHandleJellyfinAuthCommands>[0];
type HandleJellyfinListCommandsMainDeps = Parameters<typeof createHandleJellyfinListCommands>[0];
type HandleJellyfinPlayCommandMainDeps = Parameters<typeof createHandleJellyfinPlayCommand>[0];
type HandleJellyfinRemoteAnnounceCommandMainDeps = Parameters<
  typeof createHandleJellyfinRemoteAnnounceCommand
>[0];

export function createBuildHandleJellyfinAuthCommandsMainDepsHandler(
  deps: HandleJellyfinAuthCommandsMainDeps,
) {
  return (): HandleJellyfinAuthCommandsMainDeps => ({
    patchRawConfig: (patch) => deps.patchRawConfig(patch),
    authenticateWithPassword: (serverUrl, username, password, clientInfo) =>
      deps.authenticateWithPassword(serverUrl, username, password, clientInfo),
    saveStoredSession: (session) => deps.saveStoredSession(session),
    clearStoredSession: () => deps.clearStoredSession(),
    logInfo: (message: string) => deps.logInfo(message),
  });
}

export function createBuildHandleJellyfinListCommandsMainDepsHandler(
  deps: HandleJellyfinListCommandsMainDeps,
) {
  return (): HandleJellyfinListCommandsMainDeps => ({
    listJellyfinLibraries: (session, clientInfo) => deps.listJellyfinLibraries(session, clientInfo),
    listJellyfinItems: (session, clientInfo, params) =>
      deps.listJellyfinItems(session, clientInfo, params),
    listJellyfinSubtitleTracks: (session, clientInfo, itemId) =>
      deps.listJellyfinSubtitleTracks(session, clientInfo, itemId),
    writeJellyfinPreviewAuth: (responsePath, payload) =>
      deps.writeJellyfinPreviewAuth(responsePath, payload),
    logInfo: (message: string) => deps.logInfo(message),
  });
}

export function createBuildHandleJellyfinPlayCommandMainDepsHandler(
  deps: HandleJellyfinPlayCommandMainDeps,
) {
  return (): HandleJellyfinPlayCommandMainDeps => ({
    playJellyfinItemInMpv: (params) => deps.playJellyfinItemInMpv(params),
    logWarn: (message: string) => deps.logWarn(message),
  });
}

export function createBuildHandleJellyfinRemoteAnnounceCommandMainDepsHandler(
  deps: HandleJellyfinRemoteAnnounceCommandMainDeps,
) {
  return (): HandleJellyfinRemoteAnnounceCommandMainDeps => ({
    startJellyfinRemoteSession: (options) => deps.startJellyfinRemoteSession(options),
    getRemoteSession: () => deps.getRemoteSession(),
    logInfo: (message: string) => deps.logInfo(message),
    logWarn: (message: string) => deps.logWarn(message),
  });
}

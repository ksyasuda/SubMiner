import type { createOpenJellyfinSetupWindowHandler } from './jellyfin-setup-window';

type OpenJellyfinSetupWindowMainDeps = Parameters<typeof createOpenJellyfinSetupWindowHandler>[0];

export function createBuildOpenJellyfinSetupWindowMainDepsHandler(
  deps: OpenJellyfinSetupWindowMainDeps,
) {
  return (): OpenJellyfinSetupWindowMainDeps => ({
    maybeFocusExistingSetupWindow: () => deps.maybeFocusExistingSetupWindow(),
    createSetupWindow: () => deps.createSetupWindow(),
    getResolvedJellyfinConfig: () => deps.getResolvedJellyfinConfig(),
    buildSetupFormHtml: (state) => deps.buildSetupFormHtml(state),
    parseSubmissionUrl: (rawUrl: string) => deps.parseSubmissionUrl(rawUrl),
    authenticateWithPassword: (server: string, username: string, password: string, clientInfo) =>
      deps.authenticateWithPassword(server, username, password, clientInfo),
    getJellyfinClientInfo: () => deps.getJellyfinClientInfo(),
    saveStoredSession: (session) => deps.saveStoredSession(session),
    clearStoredSession: () => deps.clearStoredSession(),
    patchJellyfinConfig: (session) => deps.patchJellyfinConfig(session),
    persistAuthenticatedSession: deps.persistAuthenticatedSession
      ? (session, clientInfo) => deps.persistAuthenticatedSession?.(session, clientInfo)
      : undefined,
    logInfo: (message: string) => deps.logInfo(message),
    logError: (message: string, error: unknown) => deps.logError(message, error),
    showMpvOsd: (message: string) => deps.showMpvOsd(message),
    clearSetupWindow: () => deps.clearSetupWindow(),
    setSetupWindow: (window) => deps.setSetupWindow(window),
    registerSetupIpcHandler: deps.registerSetupIpcHandler
      ? (handler) => deps.registerSetupIpcHandler?.(handler) ?? (() => undefined)
      : undefined,
    encodeURIComponent: (value: string) => deps.encodeURIComponent(value),
    defaultServerUrl: deps.defaultServerUrl,
    hasStoredSession: () => deps.hasStoredSession(),
  });
}

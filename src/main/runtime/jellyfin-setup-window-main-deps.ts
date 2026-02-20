import type { createOpenJellyfinSetupWindowHandler } from './jellyfin-setup-window';

type OpenJellyfinSetupWindowMainDeps = Parameters<typeof createOpenJellyfinSetupWindowHandler>[0];

export function createBuildOpenJellyfinSetupWindowMainDepsHandler(
  deps: OpenJellyfinSetupWindowMainDeps,
) {
  return (): OpenJellyfinSetupWindowMainDeps => ({
    maybeFocusExistingSetupWindow: () => deps.maybeFocusExistingSetupWindow(),
    createSetupWindow: () => deps.createSetupWindow(),
    getResolvedJellyfinConfig: () => deps.getResolvedJellyfinConfig(),
    buildSetupFormHtml: (defaultServer: string, defaultUser: string) =>
      deps.buildSetupFormHtml(defaultServer, defaultUser),
    parseSubmissionUrl: (rawUrl: string) => deps.parseSubmissionUrl(rawUrl),
    authenticateWithPassword: (server: string, username: string, password: string, clientInfo) =>
      deps.authenticateWithPassword(server, username, password, clientInfo),
    getJellyfinClientInfo: () => deps.getJellyfinClientInfo(),
    saveStoredToken: (token: string) => deps.saveStoredToken(token),
    patchJellyfinConfig: (session) => deps.patchJellyfinConfig(session),
    logInfo: (message: string) => deps.logInfo(message),
    logError: (message: string, error: unknown) => deps.logError(message, error),
    showMpvOsd: (message: string) => deps.showMpvOsd(message),
    clearSetupWindow: () => deps.clearSetupWindow(),
    setSetupWindow: (window) => deps.setSetupWindow(window),
    encodeURIComponent: (value: string) => deps.encodeURIComponent(value),
  });
}

import type { createOpenAnilistSetupWindowHandler } from './anilist-setup-window';

type OpenAnilistSetupWindowMainDeps = Parameters<typeof createOpenAnilistSetupWindowHandler>[0];

export function createBuildOpenAnilistSetupWindowMainDepsHandler(
  deps: OpenAnilistSetupWindowMainDeps,
) {
  return (): OpenAnilistSetupWindowMainDeps => ({
    maybeFocusExistingSetupWindow: () => deps.maybeFocusExistingSetupWindow(),
    createSetupWindow: () => deps.createSetupWindow(),
    buildAuthorizeUrl: () => deps.buildAuthorizeUrl(),
    consumeCallbackUrl: (rawUrl: string) => deps.consumeCallbackUrl(rawUrl),
    openSetupInBrowser: (authorizeUrl: string) => deps.openSetupInBrowser(authorizeUrl),
    loadManualTokenEntry: (setupWindow, authorizeUrl: string) =>
      deps.loadManualTokenEntry(setupWindow, authorizeUrl),
    redirectUri: deps.redirectUri,
    developerSettingsUrl: deps.developerSettingsUrl,
    isAllowedExternalUrl: (url: string) => deps.isAllowedExternalUrl(url),
    isAllowedNavigationUrl: (url: string) => deps.isAllowedNavigationUrl(url),
    logWarn: (message: string, details?: unknown) => deps.logWarn(message, details),
    logError: (message: string, details: unknown) => deps.logError(message, details),
    clearSetupWindow: () => deps.clearSetupWindow(),
    setSetupPageOpened: (opened: boolean) => deps.setSetupPageOpened(opened),
    setSetupWindow: (setupWindow) => deps.setSetupWindow(setupWindow),
    openExternal: (url: string) => deps.openExternal(url),
  });
}

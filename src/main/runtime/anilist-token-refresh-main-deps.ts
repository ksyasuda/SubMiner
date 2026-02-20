import type { createRefreshAnilistClientSecretStateHandler } from './anilist-token-refresh';

type RefreshAnilistClientSecretStateMainDeps = Parameters<
  typeof createRefreshAnilistClientSecretStateHandler
>[0];

export function createBuildRefreshAnilistClientSecretStateMainDepsHandler(
  deps: RefreshAnilistClientSecretStateMainDeps,
) {
  return (): RefreshAnilistClientSecretStateMainDeps => ({
    getResolvedConfig: () => deps.getResolvedConfig(),
    isAnilistTrackingEnabled: (config) => deps.isAnilistTrackingEnabled(config),
    getCachedAccessToken: () => deps.getCachedAccessToken(),
    setCachedAccessToken: (token) => deps.setCachedAccessToken(token),
    saveStoredToken: (token: string) => deps.saveStoredToken(token),
    loadStoredToken: () => deps.loadStoredToken(),
    setClientSecretState: (state) => deps.setClientSecretState(state),
    getAnilistSetupPageOpened: () => deps.getAnilistSetupPageOpened(),
    setAnilistSetupPageOpened: (opened: boolean) => deps.setAnilistSetupPageOpened(opened),
    openAnilistSetupWindow: () => deps.openAnilistSetupWindow(),
    now: () => deps.now(),
  });
}

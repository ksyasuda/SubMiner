import type { createCriticalConfigErrorHandler, createReloadConfigHandler } from './startup-config';

type ReloadConfigMainDeps = Parameters<typeof createReloadConfigHandler>[0];
type CriticalConfigErrorMainDeps = Parameters<typeof createCriticalConfigErrorHandler>[0];

export function createBuildReloadConfigMainDepsHandler(deps: ReloadConfigMainDeps) {
  return (): ReloadConfigMainDeps => ({
    reloadConfigStrict: () => deps.reloadConfigStrict(),
    logInfo: (message: string) => deps.logInfo(message),
    logWarning: (message: string) => deps.logWarning(message),
    showDesktopNotification: (title: string, options: { body: string }) =>
      deps.showDesktopNotification(title, options),
    startConfigHotReload: () => deps.startConfigHotReload(),
    refreshAnilistClientSecretState: (options: { force: boolean }) =>
      deps.refreshAnilistClientSecretState(options),
    failHandlers: {
      logError: (details: string) => deps.failHandlers.logError(details),
      showErrorBox: (title: string, details: string) =>
        deps.failHandlers.showErrorBox(title, details),
      quit: () => deps.failHandlers.quit(),
    },
  });
}

export function createBuildCriticalConfigErrorMainDepsHandler(deps: CriticalConfigErrorMainDeps) {
  return (): CriticalConfigErrorMainDeps => ({
    getConfigPath: () => deps.getConfigPath(),
    failHandlers: {
      logError: (details: string) => deps.failHandlers.logError(details),
      showErrorBox: (title: string, details: string) =>
        deps.failHandlers.showErrorBox(title, details),
      quit: () => deps.failHandlers.quit(),
    },
  });
}

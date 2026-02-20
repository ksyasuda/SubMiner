export function createBuildReloadConfigMainDepsHandler(deps: {
  reloadConfigStrict: () => unknown;
  logInfo: (message: string) => void;
  logWarning: (message: string) => void;
  showDesktopNotification: (title: string, options: { body: string }) => void;
  startConfigHotReload: () => void;
  refreshAnilistClientSecretState: (options: { force: boolean }) => Promise<unknown>;
  failHandlers: {
    logError: (details: string) => void;
    showErrorBox: (title: string, details: string) => void;
    quit: () => void;
  };
}) {
  return () => ({
    reloadConfigStrict: () => deps.reloadConfigStrict() as never,
    logInfo: (message: string) => deps.logInfo(message),
    logWarning: (message: string) => deps.logWarning(message),
    showDesktopNotification: (title: string, options: { body: string }) =>
      deps.showDesktopNotification(title, options),
    startConfigHotReload: () => deps.startConfigHotReload(),
    refreshAnilistClientSecretState: (options: { force: boolean }) =>
      deps.refreshAnilistClientSecretState(options),
    failHandlers: {
      logError: (details: string) => deps.failHandlers.logError(details),
      showErrorBox: (title: string, details: string) => deps.failHandlers.showErrorBox(title, details),
      quit: () => deps.failHandlers.quit(),
    },
  });
}

export function createBuildCriticalConfigErrorMainDepsHandler(deps: {
  getConfigPath: () => string;
  failHandlers: {
    logError: (details: string) => void;
    showErrorBox: (title: string, details: string) => void;
    quit: () => void;
  };
}) {
  return () => ({
    getConfigPath: () => deps.getConfigPath(),
    failHandlers: {
      logError: (details: string) => deps.failHandlers.logError(details),
      showErrorBox: (title: string, details: string) => deps.failHandlers.showErrorBox(title, details),
      quit: () => deps.failHandlers.quit(),
    },
  });
}

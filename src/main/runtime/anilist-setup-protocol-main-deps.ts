import type {
  createConsumeAnilistSetupTokenFromUrlHandler,
  createHandleAnilistSetupProtocolUrlHandler,
  createNotifyAnilistSetupHandler,
  createRegisterSubminerProtocolClientHandler,
} from './anilist-setup-protocol';

type NotifyAnilistSetupMainDeps = Parameters<typeof createNotifyAnilistSetupHandler>[0];
type ConsumeAnilistSetupTokenMainDeps = Parameters<
  typeof createConsumeAnilistSetupTokenFromUrlHandler
>[0];
type HandleAnilistSetupProtocolUrlMainDeps = Parameters<
  typeof createHandleAnilistSetupProtocolUrlHandler
>[0];
type RegisterSubminerProtocolClientMainDeps = Parameters<
  typeof createRegisterSubminerProtocolClientHandler
>[0];

export function createBuildNotifyAnilistSetupMainDepsHandler(deps: NotifyAnilistSetupMainDeps) {
  return (): NotifyAnilistSetupMainDeps => ({
    hasMpvClient: () => deps.hasMpvClient(),
    showMpvOsd: (message: string) => deps.showMpvOsd(message),
    showDesktopNotification: (title: string, options: { body: string }) =>
      deps.showDesktopNotification(title, options),
    logInfo: (message: string) => deps.logInfo(message),
  });
}

export function createBuildConsumeAnilistSetupTokenFromUrlMainDepsHandler(
  deps: ConsumeAnilistSetupTokenMainDeps,
) {
  return (): ConsumeAnilistSetupTokenMainDeps => ({
    consumeAnilistSetupCallbackUrl: (input) => deps.consumeAnilistSetupCallbackUrl(input),
    saveToken: (token: string) => deps.saveToken(token),
    setCachedToken: (token: string) => deps.setCachedToken(token),
    setResolvedState: (resolvedAt: number) => deps.setResolvedState(resolvedAt),
    setSetupPageOpened: (opened: boolean) => deps.setSetupPageOpened(opened),
    onSuccess: () => deps.onSuccess(),
    closeWindow: () => deps.closeWindow(),
  });
}

export function createBuildHandleAnilistSetupProtocolUrlMainDepsHandler(
  deps: HandleAnilistSetupProtocolUrlMainDeps,
) {
  return (): HandleAnilistSetupProtocolUrlMainDeps => ({
    consumeAnilistSetupTokenFromUrl: (rawUrl: string) => deps.consumeAnilistSetupTokenFromUrl(rawUrl),
    logWarn: (message: string, details: unknown) => deps.logWarn(message, details),
  });
}

export function createBuildRegisterSubminerProtocolClientMainDepsHandler(
  deps: RegisterSubminerProtocolClientMainDeps,
) {
  return (): RegisterSubminerProtocolClientMainDeps => ({
    isDefaultApp: () => deps.isDefaultApp(),
    getArgv: () => deps.getArgv(),
    execPath: deps.execPath,
    resolvePath: (value: string) => deps.resolvePath(value),
    setAsDefaultProtocolClient: (scheme: string, path?: string, args?: string[]) =>
      deps.setAsDefaultProtocolClient(scheme, path, args),
    logWarn: (message: string, details?: unknown) => deps.logWarn(message, details),
  });
}

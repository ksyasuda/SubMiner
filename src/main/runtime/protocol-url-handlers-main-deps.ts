import type { registerProtocolUrlHandlers } from './protocol-url-handlers';

type RegisterProtocolUrlHandlersMainDeps = Parameters<typeof registerProtocolUrlHandlers>[0];

export function createBuildRegisterProtocolUrlHandlersMainDepsHandler(
  deps: RegisterProtocolUrlHandlersMainDeps,
) {
  return (): RegisterProtocolUrlHandlersMainDeps => ({
    registerOpenUrl: (listener) => deps.registerOpenUrl(listener),
    registerSecondInstance: (listener) => deps.registerSecondInstance(listener),
    handleAnilistSetupProtocolUrl: (rawUrl: string) => deps.handleAnilistSetupProtocolUrl(rawUrl),
    findAnilistSetupDeepLinkArgvUrl: (argv: string[]) => deps.findAnilistSetupDeepLinkArgvUrl(argv),
    logUnhandledOpenUrl: (rawUrl: string) => deps.logUnhandledOpenUrl(rawUrl),
    logUnhandledSecondInstanceUrl: (rawUrl: string) => deps.logUnhandledSecondInstanceUrl(rawUrl),
  });
}

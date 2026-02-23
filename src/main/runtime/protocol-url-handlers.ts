export function registerProtocolUrlHandlers(deps: {
  registerOpenUrl: (
    listener: (event: { preventDefault: () => void }, rawUrl: string) => void,
  ) => void;
  registerSecondInstance: (listener: (_event: unknown, argv: string[]) => void) => void;
  handleAnilistSetupProtocolUrl: (rawUrl: string) => boolean;
  findAnilistSetupDeepLinkArgvUrl: (argv: string[]) => string | null;
  logUnhandledOpenUrl: (rawUrl: string) => void;
  logUnhandledSecondInstanceUrl: (rawUrl: string) => void;
}) {
  deps.registerOpenUrl((event, rawUrl) => {
    event.preventDefault();
    if (!deps.handleAnilistSetupProtocolUrl(rawUrl)) {
      deps.logUnhandledOpenUrl(rawUrl);
    }
  });

  deps.registerSecondInstance((_event, argv) => {
    const rawUrl = deps.findAnilistSetupDeepLinkArgvUrl(argv);
    if (!rawUrl) {
      return;
    }
    if (!deps.handleAnilistSetupProtocolUrl(rawUrl)) {
      deps.logUnhandledSecondInstanceUrl(rawUrl);
    }
  });
}

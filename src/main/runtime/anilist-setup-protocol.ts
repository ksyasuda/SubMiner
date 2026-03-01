export type ConsumeAnilistSetupTokenDeps = {
  consumeAnilistSetupCallbackUrl: (input: {
    rawUrl: string;
    saveToken: (token: string) => void;
    setCachedToken: (token: string) => void;
    setResolvedState: (resolvedAt: number) => void;
    setSetupPageOpened: (opened: boolean) => void;
    onSuccess: () => void;
    closeWindow: () => void;
  }) => boolean;
  saveToken: (token: string) => void;
  setCachedToken: (token: string) => void;
  setResolvedState: (resolvedAt: number) => void;
  setSetupPageOpened: (opened: boolean) => void;
  onSuccess: () => void;
  closeWindow: () => void;
};

export function createConsumeAnilistSetupTokenFromUrlHandler(deps: ConsumeAnilistSetupTokenDeps) {
  return (rawUrl: string): boolean =>
    deps.consumeAnilistSetupCallbackUrl({
      rawUrl,
      saveToken: deps.saveToken,
      setCachedToken: deps.setCachedToken,
      setResolvedState: deps.setResolvedState,
      setSetupPageOpened: deps.setSetupPageOpened,
      onSuccess: deps.onSuccess,
      closeWindow: deps.closeWindow,
    });
}

export function createNotifyAnilistSetupHandler(deps: {
  hasMpvClient: () => boolean;
  showMpvOsd: (message: string) => void;
  showDesktopNotification: (title: string, options: { body: string }) => void;
  logInfo: (message: string) => void;
}) {
  return (message: string): void => {
    if (deps.hasMpvClient()) {
      deps.showMpvOsd(message);
      return;
    }
    deps.showDesktopNotification('SubMiner AniList', { body: message });
    deps.logInfo(`[AniList setup] ${message}`);
  };
}

export function createHandleAnilistSetupProtocolUrlHandler(deps: {
  consumeAnilistSetupTokenFromUrl: (rawUrl: string) => boolean;
  logWarn: (message: string, details: unknown) => void;
}) {
  return (rawUrl: string): boolean => {
    if (!rawUrl.startsWith('subminer://anilist-setup')) {
      return false;
    }
    if (deps.consumeAnilistSetupTokenFromUrl(rawUrl)) {
      return true;
    }
    deps.logWarn('AniList setup protocol URL missing access token', { rawUrl });
    return true;
  };
}

export function createRegisterSubminerProtocolClientHandler(deps: {
  isDefaultApp: () => boolean;
  getArgv: () => string[];
  execPath: string;
  resolvePath: (value: string) => string;
  setAsDefaultProtocolClient: (scheme: string, path?: string, args?: string[]) => boolean;
  logWarn: (message: string, details?: unknown) => void;
}) {
  return (): void => {
    try {
      const defaultAppEntry = deps.isDefaultApp() ? deps.getArgv()[1] : undefined;
      const success = defaultAppEntry
        ? deps.setAsDefaultProtocolClient('subminer', deps.execPath, [
            deps.resolvePath(defaultAppEntry),
          ])
        : deps.setAsDefaultProtocolClient('subminer');
      if (!success) {
        deps.logWarn('Failed to register default protocol handler for subminer:// URLs');
      }
    } catch (error) {
      deps.logWarn('Failed to register subminer:// protocol handler', error);
    }
  };
}

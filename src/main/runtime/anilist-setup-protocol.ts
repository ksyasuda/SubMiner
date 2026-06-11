import type { NotificationType, OverlayNotificationPayload } from '../../types/notification';

export type ConsumeAnilistSetupTokenDeps = {
  consumeAnilistSetupCallbackUrl: (input: {
    rawUrl: string;
    saveToken: (token: string) => boolean;
    setCachedToken: (token: string) => void;
    setResolvedState: (resolvedAt: number) => void;
    setSetupPageOpened: (opened: boolean) => void;
    onSuccess: () => void;
    closeWindow: () => void;
  }) => boolean;
  saveToken: (token: string) => boolean;
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
  getNotificationType?: () => NotificationType | undefined;
  hasMpvClient: () => boolean;
  showMpvOsd: (message: string) => void;
  showOverlayNotification?: (payload: OverlayNotificationPayload) => void;
  showDesktopNotification: (title: string, options: { body: string }) => void;
  logInfo: (message: string) => void;
}) {
  return (message: string): void => {
    const type = deps.getNotificationType?.();
    if (type) {
      if (type === 'none') {
        return;
      }
      if (type === 'overlay' || type === 'both') {
        deps.showOverlayNotification?.({
          title: 'SubMiner AniList',
          body: message,
          variant: 'success',
        });
      }
      if ((type === 'osd' || type === 'osd-system') && deps.hasMpvClient()) {
        deps.showMpvOsd(message);
      }
      if (type === 'system' || type === 'both' || type === 'osd-system') {
        deps.showDesktopNotification('SubMiner AniList', { body: message });
      }
      return;
    }

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
  logDebug: (message: string, details?: unknown) => void;
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
        deps.logDebug('Failed to register default protocol handler for subminer:// URLs');
      }
    } catch (error) {
      deps.logDebug('Failed to register subminer:// protocol handler', error);
    }
  };
}

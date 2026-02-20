type SetupWindowLike = {
  isDestroyed: () => boolean;
};

type OpenHandlerDecision = { action: 'deny' };

type FocusableWindowLike = {
  focus: () => void;
};

export function createHandleManualAnilistSetupSubmissionHandler(deps: {
  consumeCallbackUrl: (rawUrl: string) => boolean;
  redirectUri: string;
  logWarn: (message: string) => void;
}) {
  return (rawUrl: string): boolean => {
    if (!rawUrl.startsWith('subminer://anilist-setup')) {
      return false;
    }
    try {
      const parsed = new URL(rawUrl);
      const accessToken = parsed.searchParams.get('access_token')?.trim() ?? '';
      if (accessToken.length > 0) {
        return deps.consumeCallbackUrl(
          `${deps.redirectUri}#access_token=${encodeURIComponent(accessToken)}`,
        );
      }
      deps.logWarn('AniList setup submission missing access token');
      return true;
    } catch {
      deps.logWarn('AniList setup submission had invalid callback input');
      return true;
    }
  };
}

export function createMaybeFocusExistingAnilistSetupWindowHandler(deps: {
  getSetupWindow: () => FocusableWindowLike | null;
}) {
  return (): boolean => {
    const window = deps.getSetupWindow();
    if (!window) {
      return false;
    }
    window.focus();
    return true;
  };
}

export function createAnilistSetupWindowOpenHandler(deps: {
  isAllowedExternalUrl: (url: string) => boolean;
  openExternal: (url: string) => void;
  logWarn: (message: string, details?: unknown) => void;
}) {
  return ({ url }: { url: string }): OpenHandlerDecision => {
    if (!deps.isAllowedExternalUrl(url)) {
      deps.logWarn('Blocked unsafe AniList setup external URL', { url });
      return { action: 'deny' };
    }
    deps.openExternal(url);
    return { action: 'deny' };
  };
}

export function createAnilistSetupWillNavigateHandler(deps: {
  handleManualSubmission: (url: string) => boolean;
  consumeCallbackUrl: (url: string) => boolean;
  redirectUri: string;
  isAllowedNavigationUrl: (url: string) => boolean;
  logWarn: (message: string, details?: unknown) => void;
}) {
  return (params: { url: string; preventDefault: () => void }): void => {
    const { url, preventDefault } = params;
    if (deps.handleManualSubmission(url)) {
      preventDefault();
      return;
    }
    if (deps.consumeCallbackUrl(url)) {
      preventDefault();
      return;
    }
    if (url.startsWith(deps.redirectUri)) {
      preventDefault();
      return;
    }
    if (url.startsWith(`${deps.redirectUri}#`)) {
      preventDefault();
      return;
    }
    if (deps.isAllowedNavigationUrl(url)) {
      return;
    }
    preventDefault();
    deps.logWarn('Blocked unsafe AniList setup navigation URL', { url });
  };
}

export function createAnilistSetupWillRedirectHandler(deps: {
  consumeCallbackUrl: (url: string) => boolean;
}) {
  return (params: { url: string; preventDefault: () => void }): void => {
    if (deps.consumeCallbackUrl(params.url)) {
      params.preventDefault();
    }
  };
}

export function createAnilistSetupDidNavigateHandler(deps: {
  consumeCallbackUrl: (url: string) => boolean;
}) {
  return (url: string): void => {
    deps.consumeCallbackUrl(url);
  };
}

export function createAnilistSetupDidFailLoadHandler(deps: {
  onLoadFailure: (details: { errorCode: number; errorDescription: string; validatedURL: string }) => void;
}) {
  return (details: { errorCode: number; errorDescription: string; validatedURL: string }): void => {
    deps.onLoadFailure(details);
  };
}

export function createAnilistSetupDidFinishLoadHandler(deps: {
  getLoadedUrl: () => string;
  onBlankPageLoaded: () => void;
}) {
  return (): void => {
    const loadedUrl = deps.getLoadedUrl();
    if (!loadedUrl || loadedUrl === 'about:blank') {
      deps.onBlankPageLoaded();
    }
  };
}

export function createHandleAnilistSetupWindowClosedHandler(deps: {
  clearSetupWindow: () => void;
  setSetupPageOpened: (opened: boolean) => void;
}) {
  return (): void => {
    deps.clearSetupWindow();
    deps.setSetupPageOpened(false);
  };
}

export function createHandleAnilistSetupWindowOpenedHandler(deps: {
  setSetupWindow: () => void;
  setSetupPageOpened: (opened: boolean) => void;
}) {
  return (): void => {
    deps.setSetupWindow();
    deps.setSetupPageOpened(true);
  };
}

export function createAnilistSetupFallbackHandler(deps: {
  authorizeUrl: string;
  developerSettingsUrl: string;
  setupWindow: SetupWindowLike;
  openSetupInBrowser: () => void;
  loadManualTokenEntry: () => void;
  logError: (message: string, details: unknown) => void;
  logWarn: (message: string) => void;
}) {
  return {
    onLoadFailure: (details: { errorCode: number; errorDescription: string; validatedURL: string }) => {
      deps.logError('AniList setup window failed to load', details);
      deps.openSetupInBrowser();
      if (!deps.setupWindow.isDestroyed()) {
        deps.loadManualTokenEntry();
      }
    },
    onBlankPageLoaded: () => {
      deps.logWarn('AniList setup loaded a blank page; using fallback');
      deps.openSetupInBrowser();
      if (!deps.setupWindow.isDestroyed()) {
        deps.loadManualTokenEntry();
      }
    },
  };
}

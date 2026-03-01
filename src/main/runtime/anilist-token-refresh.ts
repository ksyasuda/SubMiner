type AnilistSecretResolutionState = {
  status: 'not_checked' | 'resolved' | 'error';
  source: 'none' | 'literal' | 'stored';
  message: string | null;
  resolvedAt: number | null;
  errorAt: number | null;
};

type ConfigWithAnilistToken = {
  anilist: {
    accessToken: string;
  };
};

export function createRefreshAnilistClientSecretStateHandler<
  TConfig extends ConfigWithAnilistToken,
>(deps: {
  getResolvedConfig: () => TConfig;
  isAnilistTrackingEnabled: (config: TConfig) => boolean;
  getCachedAccessToken: () => string | null;
  setCachedAccessToken: (token: string | null) => void;
  saveStoredToken: (token: string) => void;
  loadStoredToken: () => string | null | undefined;
  setClientSecretState: (state: AnilistSecretResolutionState) => void;
  getAnilistSetupPageOpened: () => boolean;
  setAnilistSetupPageOpened: (opened: boolean) => void;
  openAnilistSetupWindow: () => void;
  now: () => number;
}) {
  return async (options?: { force?: boolean }): Promise<string | null> => {
    const resolved = deps.getResolvedConfig();
    const now = deps.now();
    if (!deps.isAnilistTrackingEnabled(resolved)) {
      deps.setCachedAccessToken(null);
      deps.setClientSecretState({
        status: 'not_checked',
        source: 'none',
        message: 'anilist tracking disabled',
        resolvedAt: null,
        errorAt: null,
      });
      deps.setAnilistSetupPageOpened(false);
      return null;
    }

    const rawAccessToken = resolved.anilist.accessToken.trim();
    if (rawAccessToken.length > 0) {
      if (options?.force || rawAccessToken !== deps.getCachedAccessToken()) {
        deps.saveStoredToken(rawAccessToken);
      }
      deps.setCachedAccessToken(rawAccessToken);
      deps.setClientSecretState({
        status: 'resolved',
        source: 'literal',
        message: 'using configured anilist.accessToken',
        resolvedAt: now,
        errorAt: null,
      });
      deps.setAnilistSetupPageOpened(false);
      return rawAccessToken;
    }

    const cachedAccessToken = deps.getCachedAccessToken();
    if (!options?.force && cachedAccessToken && cachedAccessToken.length > 0) {
      return cachedAccessToken;
    }

    const storedToken = deps.loadStoredToken()?.trim() ?? '';
    if (storedToken.length > 0) {
      deps.setCachedAccessToken(storedToken);
      deps.setClientSecretState({
        status: 'resolved',
        source: 'stored',
        message: 'using stored anilist access token',
        resolvedAt: now,
        errorAt: null,
      });
      deps.setAnilistSetupPageOpened(false);
      return storedToken;
    }

    deps.setCachedAccessToken(null);
    deps.setClientSecretState({
      status: 'error',
      source: 'none',
      message: 'cannot authenticate without anilist.accessToken',
      resolvedAt: null,
      errorAt: now,
    });
    if (deps.isAnilistTrackingEnabled(resolved) && !deps.getAnilistSetupPageOpened()) {
      deps.openAnilistSetupWindow();
    }
    return null;
  };
}

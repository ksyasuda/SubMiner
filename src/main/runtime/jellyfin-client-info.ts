export function createGetResolvedJellyfinConfigHandler(deps: {
  getResolvedConfig: () => { jellyfin: unknown };
  loadStoredToken: () => string | null | undefined;
}) {
  return () => {
    const jellyfin = deps.getResolvedConfig().jellyfin as {
      accessToken?: string;
      [key: string]: unknown;
    };
    const configToken = jellyfin.accessToken?.trim() ?? '';
    if (configToken.length > 0) {
      return jellyfin as never;
    }
    const storedToken = deps.loadStoredToken()?.trim() ?? '';
    if (storedToken.length === 0) {
      return jellyfin as never;
    }
    return {
      ...jellyfin,
      accessToken: storedToken,
    } as never;
  };
}

export function createGetJellyfinClientInfoHandler(deps: {
  getResolvedJellyfinConfig: () => {
    clientName?: string;
    clientVersion?: string;
    deviceId?: string;
  };
  getDefaultJellyfinConfig: () => {
    clientName?: string;
    clientVersion?: string;
    deviceId?: string;
  };
}) {
  return (
    config = deps.getResolvedJellyfinConfig(),
  ): {
    clientName: string;
    clientVersion: string;
    deviceId: string;
  } => {
    const defaults = deps.getDefaultJellyfinConfig();
    return {
      clientName: config.clientName || defaults.clientName || '',
      clientVersion: config.clientVersion || defaults.clientVersion || '',
      deviceId: config.deviceId || defaults.deviceId || '',
    };
  };
}

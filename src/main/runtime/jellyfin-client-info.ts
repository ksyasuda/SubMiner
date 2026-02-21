export function createGetResolvedJellyfinConfigHandler(deps: {
  getResolvedConfig: () => { jellyfin: unknown };
  loadStoredSession: () => { accessToken: string; userId: string } | null | undefined;
  getEnv: (name: string) => string | undefined;
}) {
  return () => {
    const jellyfin = deps.getResolvedConfig().jellyfin as {
      userId?: string;
      [key: string]: unknown;
    };

    const envToken = deps.getEnv('SUBMINER_JELLYFIN_ACCESS_TOKEN')?.trim() ?? '';
    const envUserId = deps.getEnv('SUBMINER_JELLYFIN_USER_ID')?.trim() ?? '';
    const stored = deps.loadStoredSession();
    const storedToken = stored?.accessToken?.trim() ?? '';
    const storedUserId = stored?.userId?.trim() ?? '';

    if (envToken.length > 0) {
      return {
        ...jellyfin,
        accessToken: envToken,
        userId: envUserId || storedUserId || '',
      } as never;
    }

    if (storedToken.length > 0 && storedUserId.length > 0) {
      return {
        ...jellyfin,
        accessToken: storedToken,
        userId: storedUserId,
      } as never;
    }

    return jellyfin as never;
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

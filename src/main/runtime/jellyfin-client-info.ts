export function createGetResolvedJellyfinConfigHandler(deps: {
  getResolvedConfig: () => { jellyfin: unknown };
}) {
  return () => deps.getResolvedConfig().jellyfin as never;
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

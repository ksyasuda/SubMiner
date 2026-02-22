import type { JellyfinStoredSession } from '../../core/services/jellyfin-token-store';
import type { ResolvedConfig } from '../../types';

type ResolvedJellyfinConfig = ResolvedConfig['jellyfin'];
type ResolvedJellyfinConfigWithSession = ResolvedJellyfinConfig & {
  accessToken?: string;
  userId?: string;
};

export function createGetResolvedJellyfinConfigHandler(deps: {
  getResolvedConfig: () => { jellyfin: ResolvedJellyfinConfig };
  loadStoredSession: () => JellyfinStoredSession | null | undefined;
  getEnv: (name: string) => string | undefined;
}) {
  return (): ResolvedJellyfinConfigWithSession => {
    const jellyfin = deps.getResolvedConfig().jellyfin;

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
      };
    }

    if (storedToken.length > 0 && storedUserId.length > 0) {
      return {
        ...jellyfin,
        accessToken: storedToken,
        userId: storedUserId,
      };
    }

    return jellyfin;
  };
}

export function createGetJellyfinClientInfoHandler(deps: {
  getResolvedJellyfinConfig: () => Partial<
    Pick<ResolvedJellyfinConfig, 'clientName' | 'clientVersion' | 'deviceId'>
  >;
  getDefaultJellyfinConfig: () => Partial<
    Pick<ResolvedJellyfinConfig, 'clientName' | 'clientVersion' | 'deviceId'>
  >;
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

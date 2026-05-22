import type { JellyfinStoredSession } from '../../core/services/jellyfin-token-store';
import type { ResolvedConfig } from '../../types';
import {
  DEFAULT_JELLYFIN_CLIENT_NAME,
  DEFAULT_JELLYFIN_CLIENT_VERSION,
  createHostDerivedJellyfinDeviceId,
} from './jellyfin-device-identity';

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
  getResolvedJellyfinConfig: () => unknown;
  getHostName?: () => string;
  defaultClientName?: string;
  defaultClientVersion?: string;
}) {
  return (
    _config = deps.getResolvedJellyfinConfig(),
  ): {
    clientName: string;
    clientVersion: string;
    deviceId: string;
  } => {
    return {
      clientName: deps.defaultClientName || DEFAULT_JELLYFIN_CLIENT_NAME,
      clientVersion: deps.defaultClientVersion || DEFAULT_JELLYFIN_CLIENT_VERSION,
      deviceId: createHostDerivedJellyfinDeviceId(deps.getHostName?.() || ''),
    };
  };
}

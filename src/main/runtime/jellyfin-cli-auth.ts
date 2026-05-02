import type { CliArgs } from '../../cli/args';

type JellyfinConfig = {
  serverUrl: string;
  username: string;
  recentServers?: string[];
};

type JellyfinClientInfo = {
  deviceId: string;
  clientName: string;
  clientVersion: string;
};

type JellyfinSession = {
  serverUrl: string;
  username: string;
  accessToken: string;
  userId: string;
};

const MAX_RECENT_JELLYFIN_SERVERS = 5;

export function normalizeJellyfinServerUrl(value: string): string {
  return value.trim().replace(/\/+$/, '');
}

export function normalizeJellyfinRecentServers(values: unknown[]): string[] {
  const seen = new Set<string>();
  const servers: string[] = [];
  for (const value of values) {
    if (typeof value !== 'string') continue;
    const normalized = normalizeJellyfinServerUrl(value);
    if (!normalized || seen.has(normalized)) continue;
    seen.add(normalized);
    servers.push(normalized);
    if (servers.length >= MAX_RECENT_JELLYFIN_SERVERS) break;
  }
  return servers;
}

export function mergeJellyfinRecentServers(serverUrl: string, existing: unknown[]): string[] {
  return normalizeJellyfinRecentServers([serverUrl, ...existing]);
}

export function persistJellyfinAuthSession(deps: {
  session: JellyfinSession;
  clientInfo: JellyfinClientInfo;
  existingRecentServers?: unknown[];
  saveStoredSession: (session: { accessToken: string; userId: string }) => void;
  patchRawConfig: (patch: {
    jellyfin: Partial<{
      enabled: boolean;
      serverUrl: string;
      username: string;
      deviceId: string;
      clientName: string;
      clientVersion: string;
      recentServers: string[];
    }>;
  }) => void;
}): void {
  deps.saveStoredSession({
    accessToken: deps.session.accessToken,
    userId: deps.session.userId,
  });
  deps.patchRawConfig({
    jellyfin: {
      enabled: true,
      serverUrl: deps.session.serverUrl,
      username: deps.session.username,
      deviceId: deps.clientInfo.deviceId,
      clientName: deps.clientInfo.clientName,
      clientVersion: deps.clientInfo.clientVersion,
      recentServers: mergeJellyfinRecentServers(
        deps.session.serverUrl,
        deps.existingRecentServers || [],
      ),
    },
  });
}

export function createHandleJellyfinAuthCommands(deps: {
  patchRawConfig: (patch: {
    jellyfin: Partial<{
      enabled: boolean;
      serverUrl: string;
      username: string;
      deviceId: string;
      clientName: string;
      clientVersion: string;
    }>;
  }) => void;
  authenticateWithPassword: (
    serverUrl: string,
    username: string,
    password: string,
    clientInfo: JellyfinClientInfo,
  ) => Promise<JellyfinSession>;
  saveStoredSession: (session: { accessToken: string; userId: string }) => void;
  clearStoredSession: () => void;
  logInfo: (message: string) => void;
}) {
  return async (params: {
    args: CliArgs;
    jellyfinConfig: JellyfinConfig;
    serverUrl: string;
    clientInfo: JellyfinClientInfo;
  }): Promise<boolean> => {
    if (params.args.jellyfinLogout) {
      deps.clearStoredSession();
      deps.patchRawConfig({
        jellyfin: {},
      });
      deps.logInfo('Cleared stored Jellyfin auth session.');
      return true;
    }

    if (!params.args.jellyfinLogin) {
      return false;
    }

    const username = (params.args.jellyfinUsername || params.jellyfinConfig.username).trim();
    const password = params.args.jellyfinPassword || '';
    const session = await deps.authenticateWithPassword(
      params.serverUrl,
      username,
      password,
      params.clientInfo,
    );
    persistJellyfinAuthSession({
      session,
      clientInfo: params.clientInfo,
      existingRecentServers: params.jellyfinConfig.recentServers || [],
      saveStoredSession: (storedSession) => deps.saveStoredSession(storedSession),
      patchRawConfig: (patch) => deps.patchRawConfig(patch),
    });
    deps.logInfo(`Jellyfin login succeeded for ${session.username}.`);
    return true;
  };
}

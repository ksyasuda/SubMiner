import type { CliArgs } from '../../cli/args';

type JellyfinConfig = {
  serverUrl: string;
  username: string;
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
    deps.saveStoredSession({
      accessToken: session.accessToken,
      userId: session.userId,
    });
    deps.patchRawConfig({
      jellyfin: {
        enabled: true,
        serverUrl: session.serverUrl,
        username: session.username,
        deviceId: params.clientInfo.deviceId,
        clientName: params.clientInfo.clientName,
        clientVersion: params.clientInfo.clientVersion,
      },
    });
    deps.logInfo(`Jellyfin login succeeded for ${session.username}.`);
    return true;
  };
}

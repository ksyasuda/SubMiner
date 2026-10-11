import type { CliArgs } from '../../cli/args';
import type { JellyfinCliResponse } from '../../shared/jellyfin-cli-response';

type JellyfinConfigBase = {
  serverUrl?: string;
  accessToken?: string;
  userId?: string;
  username?: string;
};

type JellyfinSession = {
  serverUrl: string;
  accessToken: string;
  userId: string;
  username: string;
};

export function createRunJellyfinCommandHandler<
  TClientInfo,
  TConfig extends JellyfinConfigBase,
>(deps: {
  getJellyfinConfig: () => TConfig;
  defaultServerUrl: string;
  getJellyfinClientInfo: (config: TConfig) => TClientInfo;
  handleAuthCommands: (params: {
    args: CliArgs;
    jellyfinConfig: TConfig;
    serverUrl: string;
    clientInfo: TClientInfo;
  }) => Promise<boolean>;
  handleRemoteAnnounceCommand: (args: CliArgs) => Promise<boolean>;
  handleListCommands: (params: {
    args: CliArgs;
    session: JellyfinSession;
    clientInfo: TClientInfo;
    jellyfinConfig: TConfig;
  }) => Promise<boolean>;
  handlePlayCommand: (params: {
    args: CliArgs;
    session: JellyfinSession;
    clientInfo: TClientInfo;
    jellyfinConfig: TConfig;
  }) => Promise<boolean>;
  writeJellyfinResponse: (responsePath: string, response: JellyfinCliResponse) => void;
}) {
  const run = async (args: CliArgs): Promise<void> => {
    const jellyfinConfig = deps.getJellyfinConfig();
    const serverUrl =
      args.jellyfinServer?.trim() || jellyfinConfig.serverUrl || deps.defaultServerUrl;
    const clientInfo = deps.getJellyfinClientInfo(jellyfinConfig);

    if (
      await deps.handleAuthCommands({
        args,
        jellyfinConfig,
        serverUrl,
        clientInfo,
      })
    ) {
      return;
    }

    const accessToken = jellyfinConfig.accessToken;
    const userId = jellyfinConfig.userId;
    if (!serverUrl || !accessToken || !userId) {
      throw new Error('Missing Jellyfin session. Run --jellyfin-login first.');
    }

    const session: JellyfinSession = {
      serverUrl,
      accessToken,
      userId,
      username: jellyfinConfig.username || '',
    };

    if (await deps.handleRemoteAnnounceCommand(args)) {
      return;
    }

    if (
      await deps.handleListCommands({
        args,
        session,
        clientInfo,
        jellyfinConfig,
      })
    ) {
      return;
    }

    if (
      await deps.handlePlayCommand({
        args,
        session,
        clientInfo,
        jellyfinConfig,
      })
    ) {
      return;
    }
  };

  // A launcher waiting on --jellyfin-response-path gets the failure reason instead of a timeout.
  return async (args: CliArgs): Promise<void> => {
    try {
      await run(args);
    } catch (error) {
      const responsePath = args.jellyfinResponsePath?.trim();
      if (responsePath) {
        try {
          deps.writeJellyfinResponse(responsePath, {
            error: error instanceof Error ? error.message : String(error),
          });
        } catch {
          // Keep the original failure; the launcher falls back to its timeout.
        }
      }
      throw error;
    }
  };
}

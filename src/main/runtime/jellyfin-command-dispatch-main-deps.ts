import type { CliArgs } from '../../cli/args';

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

export type RunJellyfinCommandMainDeps<TClientInfo, TConfig extends JellyfinConfigBase> = {
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
};

export function createBuildRunJellyfinCommandMainDepsHandler<
  TClientInfo,
  TConfig extends JellyfinConfigBase,
>(deps: RunJellyfinCommandMainDeps<TClientInfo, TConfig>) {
  return (): RunJellyfinCommandMainDeps<TClientInfo, TConfig> => ({
    getJellyfinConfig: () => deps.getJellyfinConfig(),
    defaultServerUrl: deps.defaultServerUrl,
    getJellyfinClientInfo: (config: TConfig) => deps.getJellyfinClientInfo(config),
    handleAuthCommands: (params) => deps.handleAuthCommands(params),
    handleRemoteAnnounceCommand: (args: CliArgs) => deps.handleRemoteAnnounceCommand(args),
    handleListCommands: (params) => deps.handleListCommands(params),
    handlePlayCommand: (params) => deps.handlePlayCommand(params),
  });
}

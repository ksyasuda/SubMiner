import type { CliArgs } from '../../cli/args';

type JellyfinSession = {
  serverUrl: string;
  accessToken: string;
  userId: string;
  username: string;
};

type JellyfinClientInfo = {
  clientName: string;
  clientVersion: string;
  deviceId: string;
};

export function createHandleJellyfinPlayCommand(deps: {
  playJellyfinItemInMpv: (params: {
    session: JellyfinSession;
    clientInfo: JellyfinClientInfo;
    jellyfinConfig: unknown;
    itemId: string;
    audioStreamIndex?: number;
    subtitleStreamIndex?: number;
    setQuitOnDisconnectArm?: boolean;
  }) => Promise<void>;
  logWarn: (message: string) => void;
}) {
  return async (params: {
    args: CliArgs;
    session: JellyfinSession;
    clientInfo: JellyfinClientInfo;
    jellyfinConfig: unknown;
  }): Promise<boolean> => {
    const { args, session, clientInfo, jellyfinConfig } = params;
    if (!args.jellyfinPlay) {
      return false;
    }
    if (!args.jellyfinItemId) {
      deps.logWarn('Ignoring --jellyfin-play without --jellyfin-item-id.');
      return true;
    }
    await deps.playJellyfinItemInMpv({
      session,
      clientInfo,
      jellyfinConfig,
      itemId: args.jellyfinItemId,
      audioStreamIndex: args.jellyfinAudioStreamIndex,
      subtitleStreamIndex: args.jellyfinSubtitleStreamIndex,
      setQuitOnDisconnectArm: true,
    });
    return true;
  };
}

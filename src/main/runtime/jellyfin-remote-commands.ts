export type ActiveJellyfinRemotePlaybackState = {
  itemId: string;
  mediaSourceId?: string;
  audioStreamIndex?: number | null;
  subtitleStreamIndex?: number | null;
  playMethod: 'DirectPlay' | 'Transcode';
};

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

type JellyfinConfigLike = {
  serverUrl?: string;
  accessToken?: string;
  userId?: string;
  username?: string;
};

function asInteger(value: unknown): number | undefined {
  if (typeof value !== 'number' || !Number.isInteger(value)) return undefined;
  return value;
}

export function getConfiguredJellyfinSession(config: JellyfinConfigLike): JellyfinSession | null {
  if (!config.serverUrl || !config.accessToken || !config.userId) {
    return null;
  }
  return {
    serverUrl: config.serverUrl,
    accessToken: config.accessToken,
    userId: config.userId,
    username: config.username || '',
  };
}

export type JellyfinRemotePlayHandlerDeps = {
  getConfiguredSession: () => JellyfinSession | null;
  getClientInfo: () => JellyfinClientInfo;
  getJellyfinConfig: () => unknown;
  playJellyfinItem: (params: {
    session: JellyfinSession;
    clientInfo: JellyfinClientInfo;
    jellyfinConfig: unknown;
    itemId: string;
    audioStreamIndex?: number;
    subtitleStreamIndex?: number;
    startTimeTicksOverride?: number;
    setQuitOnDisconnectArm?: boolean;
  }) => Promise<void>;
  logWarn: (message: string) => void;
};

export function createHandleJellyfinRemotePlay(deps: JellyfinRemotePlayHandlerDeps) {
  return async (payload: unknown): Promise<void> => {
    const session = deps.getConfiguredSession();
    if (!session) return;
    const clientInfo = deps.getClientInfo();
    const jellyfinConfig = deps.getJellyfinConfig();
    const data = payload && typeof payload === 'object' ? (payload as Record<string, unknown>) : {};
    const itemIds = Array.isArray(data.ItemIds)
      ? data.ItemIds.filter((entry): entry is string => typeof entry === 'string')
      : [];
    const itemId = itemIds[0];
    if (!itemId) {
      deps.logWarn('Ignoring Jellyfin remote Play event without ItemIds.');
      return;
    }
    await deps.playJellyfinItem({
      session,
      clientInfo,
      jellyfinConfig,
      itemId,
      audioStreamIndex: asInteger(data.AudioStreamIndex),
      subtitleStreamIndex: asInteger(data.SubtitleStreamIndex),
      startTimeTicksOverride: asInteger(data.StartPositionTicks),
      setQuitOnDisconnectArm: false,
    });
  };
}

type MpvClientLike = object;

export type JellyfinRemotePlaystateHandlerDeps = {
  getMpvClient: () => MpvClientLike | null;
  sendMpvCommand: (client: MpvClientLike, command: (string | number)[]) => void;
  reportJellyfinRemoteProgress: (force: boolean) => Promise<void>;
  reportJellyfinRemoteStopped: () => Promise<void>;
  jellyfinTicksToSeconds: (ticks: number) => number;
};

export function createHandleJellyfinRemotePlaystate(deps: JellyfinRemotePlaystateHandlerDeps) {
  return async (payload: unknown): Promise<void> => {
    const data = payload && typeof payload === 'object' ? (payload as Record<string, unknown>) : {};
    const command = String(data.Command || '');
    const client = deps.getMpvClient();
    if (!client) return;
    if (command === 'Pause') {
      deps.sendMpvCommand(client, ['set_property', 'pause', 'yes']);
      await deps.reportJellyfinRemoteProgress(true);
      return;
    }
    if (command === 'Unpause') {
      deps.sendMpvCommand(client, ['set_property', 'pause', 'no']);
      await deps.reportJellyfinRemoteProgress(true);
      return;
    }
    if (command === 'PlayPause') {
      deps.sendMpvCommand(client, ['cycle', 'pause']);
      await deps.reportJellyfinRemoteProgress(true);
      return;
    }
    if (command === 'Stop') {
      deps.sendMpvCommand(client, ['stop']);
      await deps.reportJellyfinRemoteStopped();
      return;
    }
    if (command === 'Seek') {
      const seekTicks = asInteger(data.SeekPositionTicks);
      if (seekTicks !== undefined) {
        deps.sendMpvCommand(client, [
          'seek',
          deps.jellyfinTicksToSeconds(seekTicks),
          'absolute+exact',
        ]);
        await deps.reportJellyfinRemoteProgress(true);
      }
    }
  };
}

export type JellyfinRemoteGeneralCommandHandlerDeps = {
  getMpvClient: () => MpvClientLike | null;
  sendMpvCommand: (client: MpvClientLike, command: (string | number)[]) => void;
  getActivePlayback: () => ActiveJellyfinRemotePlaybackState | null;
  reportJellyfinRemoteProgress: (force: boolean) => Promise<void>;
  logDebug: (message: string) => void;
};

export function createHandleJellyfinRemoteGeneralCommand(
  deps: JellyfinRemoteGeneralCommandHandlerDeps,
) {
  return async (payload: unknown): Promise<void> => {
    const data = payload && typeof payload === 'object' ? (payload as Record<string, unknown>) : {};
    const command = String(data.Name || '');
    const args =
      data.Arguments && typeof data.Arguments === 'object'
        ? (data.Arguments as Record<string, unknown>)
        : {};
    const client = deps.getMpvClient();
    if (!client) return;

    if (command === 'SetAudioStreamIndex') {
      const index = asInteger(args.Index);
      if (index !== undefined) {
        deps.sendMpvCommand(client, ['set_property', 'aid', index]);
        const playback = deps.getActivePlayback();
        if (playback) {
          playback.audioStreamIndex = index;
        }
        await deps.reportJellyfinRemoteProgress(true);
      }
      return;
    }
    if (command === 'SetSubtitleStreamIndex') {
      const index = asInteger(args.Index);
      if (index !== undefined) {
        deps.sendMpvCommand(client, ['set_property', 'sid', index < 0 ? 'no' : index]);
        const playback = deps.getActivePlayback();
        if (playback) {
          playback.subtitleStreamIndex = index < 0 ? null : index;
        }
        await deps.reportJellyfinRemoteProgress(true);
      }
      return;
    }

    deps.logDebug(`Ignoring unsupported Jellyfin GeneralCommand: ${command}`);
  };
}

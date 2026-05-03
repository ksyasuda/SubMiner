type JellyfinRemoteConfig = {
  enabled: boolean;
  remoteControlEnabled: boolean;
  remoteControlAutoConnect: boolean;
  serverUrl: string;
  accessToken?: string;
  userId?: string;
  deviceId: string;
  clientName: string;
  clientVersion: string;
  remoteControlDeviceName: string;
  autoAnnounce: boolean;
};

type JellyfinRemoteService = {
  start: () => void;
  stop: () => void;
  advertiseNow: () => Promise<boolean>;
};

type JellyfinRemoteEventPayload = unknown;

type JellyfinRemoteServiceOptions = {
  serverUrl: string;
  accessToken: string;
  deviceId: string;
  clientName: string;
  clientVersion: string;
  deviceName: string;
  capabilities: {
    PlayableMediaTypes: string;
    SupportedCommands: string;
    SupportsMediaControl: boolean;
  };
  onConnected: () => void;
  onDisconnected: () => void;
  onPlay: (payload: JellyfinRemoteEventPayload) => void;
  onPlaystate: (payload: JellyfinRemoteEventPayload) => void;
  onGeneralCommand: (payload: JellyfinRemoteEventPayload) => void;
};

export function createStartJellyfinRemoteSessionHandler(deps: {
  getJellyfinConfig: () => JellyfinRemoteConfig;
  getCurrentSession: () => JellyfinRemoteService | null;
  setCurrentSession: (session: JellyfinRemoteService | null) => void;
  createRemoteSessionService: (options: JellyfinRemoteServiceOptions) => JellyfinRemoteService;
  defaultDeviceId: string;
  defaultClientName: string;
  defaultClientVersion: string;
  handlePlay: (payload: JellyfinRemoteEventPayload) => Promise<void>;
  handlePlaystate: (payload: JellyfinRemoteEventPayload) => Promise<void>;
  handleGeneralCommand: (payload: JellyfinRemoteEventPayload) => Promise<void>;
  logInfo: (message: string) => void;
  logWarn: (message: string, details?: unknown) => void;
}) {
  return async (options?: { explicit?: boolean }): Promise<void> => {
    const jellyfinConfig = deps.getJellyfinConfig();
    if (jellyfinConfig.enabled === false) return;
    if (jellyfinConfig.remoteControlEnabled === false) return;
    if (jellyfinConfig.remoteControlAutoConnect === false && options?.explicit !== true) return;
    if (!jellyfinConfig.serverUrl || !jellyfinConfig.accessToken || !jellyfinConfig.userId) return;

    const existing = deps.getCurrentSession();
    if (existing) {
      existing.stop();
      deps.setCurrentSession(null);
    }

    const service = deps.createRemoteSessionService({
      serverUrl: jellyfinConfig.serverUrl,
      accessToken: jellyfinConfig.accessToken,
      deviceId: jellyfinConfig.deviceId || deps.defaultDeviceId,
      clientName: jellyfinConfig.clientName || deps.defaultClientName,
      clientVersion: jellyfinConfig.clientVersion || deps.defaultClientVersion,
      deviceName:
        jellyfinConfig.remoteControlDeviceName ||
        jellyfinConfig.clientName ||
        deps.defaultClientName,
      capabilities: {
        PlayableMediaTypes: 'Video,Audio',
        SupportedCommands:
          'Play,Playstate,PlayMediaSource,SetAudioStreamIndex,SetSubtitleStreamIndex,Mute,Unmute,SetVolume,DisplayContent',
        SupportsMediaControl: true,
      },
      onConnected: () => {
        deps.logInfo('Jellyfin remote websocket connected.');
        if (jellyfinConfig.autoAnnounce) {
          void service.advertiseNow().then((registered) => {
            if (registered) {
              deps.logInfo('Jellyfin cast target is visible to server sessions.');
            } else {
              deps.logWarn(
                'Jellyfin remote connected but device not visible in server sessions yet.',
              );
            }
          });
        }
      },
      onDisconnected: () => {
        deps.logWarn('Jellyfin remote websocket disconnected; retrying.');
      },
      onPlay: (payload) => {
        void deps.handlePlay(payload).catch((error) => {
          deps.logWarn('Failed handling Jellyfin remote Play event', error);
        });
      },
      onPlaystate: (payload) => {
        void deps.handlePlaystate(payload).catch((error) => {
          deps.logWarn('Failed handling Jellyfin remote Playstate event', error);
        });
      },
      onGeneralCommand: (payload) => {
        void deps.handleGeneralCommand(payload).catch((error) => {
          deps.logWarn('Failed handling Jellyfin remote GeneralCommand event', error);
        });
      },
    });

    service.start();
    deps.setCurrentSession(service);
    deps.logInfo(
      `Jellyfin remote session enabled (${jellyfinConfig.remoteControlDeviceName || jellyfinConfig.clientName || 'SubMiner'}).`,
    );
  };
}

export function createStopJellyfinRemoteSessionHandler(deps: {
  getCurrentSession: () => JellyfinRemoteService | null;
  setCurrentSession: (session: JellyfinRemoteService | null) => void;
  clearActivePlayback: () => void;
}) {
  return (): void => {
    const session = deps.getCurrentSession();
    if (!session) return;
    session.stop();
    deps.setCurrentSession(null);
    deps.clearActivePlayback();
  };
}

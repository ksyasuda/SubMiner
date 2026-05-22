import { resolveJellyfinRemoteDeviceName } from './jellyfin-device-identity';

type JellyfinRemoteConfig = {
  enabled: boolean;
  remoteControlEnabled: boolean;
  remoteControlAutoConnect: boolean;
  serverUrl: string;
  accessToken?: string;
  userId?: string;
  autoAnnounce: boolean;
};

type JellyfinClientInfo = {
  deviceId: string;
  clientName: string;
  clientVersion: string;
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
  getClientInfo: () => JellyfinClientInfo;
  getHostName: () => string;
  defaultDeviceId: string;
  defaultClientName: string;
  defaultClientVersion: string;
  handlePlay: (payload: JellyfinRemoteEventPayload) => Promise<void>;
  handlePlaystate: (payload: JellyfinRemoteEventPayload) => Promise<void>;
  handleGeneralCommand: (payload: JellyfinRemoteEventPayload) => Promise<void>;
  logInfo: (message: string) => void;
  logWarn: (message: string, details?: unknown) => void;
  onSessionStateChanged?: () => void;
}) {
  return async (options?: { explicit?: boolean }): Promise<void> => {
    const jellyfinConfig = deps.getJellyfinConfig();
    if (jellyfinConfig.enabled === false) return;
    if (jellyfinConfig.remoteControlEnabled === false) return;
    if (jellyfinConfig.remoteControlAutoConnect === false && options?.explicit !== true) return;
    if (!jellyfinConfig.serverUrl || !jellyfinConfig.accessToken || !jellyfinConfig.userId) return;

    const clientInfo = deps.getClientInfo();
    const clientName = clientInfo.clientName || deps.defaultClientName;
    const clientVersion = clientInfo.clientVersion || deps.defaultClientVersion;
    const deviceName = resolveJellyfinRemoteDeviceName({
      hostName: deps.getHostName(),
    });

    const existing = deps.getCurrentSession();
    if (existing) {
      existing.stop();
      deps.setCurrentSession(null);
    }

    const service = deps.createRemoteSessionService({
      serverUrl: jellyfinConfig.serverUrl,
      accessToken: jellyfinConfig.accessToken,
      deviceId: clientInfo.deviceId || deps.defaultDeviceId,
      clientName,
      clientVersion,
      deviceName,
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
    deps.onSessionStateChanged?.();
    deps.logInfo(`Jellyfin remote session enabled (${deviceName}).`);
  };
}

export function createStopJellyfinRemoteSessionHandler(deps: {
  getCurrentSession: () => JellyfinRemoteService | null;
  setCurrentSession: (session: JellyfinRemoteService | null) => void;
  clearActivePlayback: () => void;
  onSessionStateChanged?: () => void;
}) {
  return (): void => {
    const session = deps.getCurrentSession();
    if (!session) return;
    session.stop();
    deps.setCurrentSession(null);
    deps.clearActivePlayback();
    deps.onSessionStateChanged?.();
  };
}

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

type JellyfinPlaybackPlan = {
  url: string;
  mode: 'direct' | 'transcode';
  title: string;
  startTimeTicks: number;
  audioStreamIndex?: number | null;
  subtitleStreamIndex?: number | null;
};

type ActivePlaybackState = {
  itemId: string;
  mediaSourceId: undefined;
  audioStreamIndex?: number | null;
  subtitleStreamIndex?: number | null;
  playMethod: 'DirectPlay' | 'Transcode';
};

type MpvClientLike = unknown;

export function createPlayJellyfinItemInMpvHandler(deps: {
  ensureMpvConnectedForPlayback: () => Promise<boolean>;
  getMpvClient: () => MpvClientLike | null;
  resolvePlaybackPlan: (params: {
    session: JellyfinSession;
    clientInfo: JellyfinClientInfo;
    jellyfinConfig: unknown;
    itemId: string;
    audioStreamIndex?: number | null;
    subtitleStreamIndex?: number | null;
  }) => Promise<JellyfinPlaybackPlan>;
  applyJellyfinMpvDefaults: (mpvClient: MpvClientLike) => void;
  sendMpvCommand: (command: Array<string | number>) => void;
  armQuitOnDisconnect: () => void;
  schedule: (callback: () => void, delayMs: number) => void;
  convertTicksToSeconds: (ticks: number) => number;
  preloadExternalSubtitles: (params: {
    session: JellyfinSession;
    clientInfo: JellyfinClientInfo;
    itemId: string;
  }) => void;
  setActivePlayback: (state: ActivePlaybackState) => void;
  setLastProgressAtMs: (value: number) => void;
  reportPlaying: (payload: {
    itemId: string;
    mediaSourceId: undefined;
    playMethod: 'DirectPlay' | 'Transcode';
    audioStreamIndex?: number | null;
    subtitleStreamIndex?: number | null;
    eventName: 'start';
  }) => void;
  showMpvOsd: (text: string) => void;
}) {
  return async (params: {
    session: JellyfinSession;
    clientInfo: JellyfinClientInfo;
    jellyfinConfig: unknown;
    itemId: string;
    audioStreamIndex?: number | null;
    subtitleStreamIndex?: number | null;
    startTimeTicksOverride?: number;
    setQuitOnDisconnectArm?: boolean;
  }): Promise<void> => {
    const connected = await deps.ensureMpvConnectedForPlayback();
    const mpvClient = deps.getMpvClient();
    if (!connected || !mpvClient) {
      throw new Error(
        'MPV not connected and auto-launch failed. Ensure mpv is installed and available in PATH.',
      );
    }

    const plan = await deps.resolvePlaybackPlan({
      session: params.session,
      clientInfo: params.clientInfo,
      jellyfinConfig: params.jellyfinConfig,
      itemId: params.itemId,
      audioStreamIndex: params.audioStreamIndex,
      subtitleStreamIndex: params.subtitleStreamIndex,
    });

    deps.applyJellyfinMpvDefaults(mpvClient);
    deps.sendMpvCommand(['set_property', 'sub-auto', 'no']);
    deps.sendMpvCommand(['loadfile', plan.url, 'replace']);
    if (params.setQuitOnDisconnectArm !== false) {
      deps.armQuitOnDisconnect();
    }
    deps.sendMpvCommand(['set_property', 'force-media-title', `[Jellyfin/${plan.mode}] ${plan.title}`]);
    deps.sendMpvCommand(['set_property', 'sid', 'no']);
    deps.schedule(() => {
      deps.sendMpvCommand(['set_property', 'sid', 'no']);
    }, 500);

    const startTimeTicks =
      typeof params.startTimeTicksOverride === 'number'
        ? Math.max(0, params.startTimeTicksOverride)
        : plan.startTimeTicks;
    if (startTimeTicks > 0) {
      deps.sendMpvCommand(['seek', deps.convertTicksToSeconds(startTimeTicks), 'absolute+exact']);
    }

    deps.preloadExternalSubtitles({
      session: params.session,
      clientInfo: params.clientInfo,
      itemId: params.itemId,
    });

    const playMethod = plan.mode === 'direct' ? 'DirectPlay' : 'Transcode';
    deps.setActivePlayback({
      itemId: params.itemId,
      mediaSourceId: undefined,
      audioStreamIndex: plan.audioStreamIndex,
      subtitleStreamIndex: plan.subtitleStreamIndex,
      playMethod,
    });
    deps.setLastProgressAtMs(0);
    deps.reportPlaying({
      itemId: params.itemId,
      mediaSourceId: undefined,
      playMethod,
      audioStreamIndex: plan.audioStreamIndex,
      subtitleStreamIndex: plan.subtitleStreamIndex,
      eventName: 'start',
    });
    deps.showMpvOsd(`Jellyfin ${plan.mode}: ${plan.title}`);
  };
}

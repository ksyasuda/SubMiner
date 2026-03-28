type DiscordPresenceServiceLike = {
  publish: (snapshot: {
    mediaTitle: string | null;
    mediaPath: string | null;
    subtitleText: string;
    currentTimeSec: number | null;
    mediaDurationSec: number | null;
    paused: boolean | null;
    connected: boolean;
    sessionStartedAtMs: number;
  }) => void;
};

type MpvClientLike = {
  connected?: boolean;
  currentTimePos?: number | null;
  requestProperty: (name: string) => Promise<unknown>;
};

export type DiscordPresenceRuntimeDeps = {
  getDiscordPresenceService: () => DiscordPresenceServiceLike | null;
  isDiscordPresenceEnabled: () => boolean;
  getMpvClient: () => MpvClientLike | null;
  getCurrentMediaTitle: () => string | null;
  getCurrentMediaPath: () => string | null;
  getCurrentSubtitleText: () => string;
  getPlaybackPaused: () => boolean | null;
  getFallbackMediaDurationSec: () => number | null;
  getSessionStartedAtMs: () => number;
  getMediaDurationSec: () => number | null;
  setMediaDurationSec: (durationSec: number | null) => void;
};

export function createDiscordPresenceRuntime(deps: DiscordPresenceRuntimeDeps) {
  const refreshDiscordPresenceMediaDuration = async (): Promise<void> => {
    const client = deps.getMpvClient();
    if (!client?.connected) {
      return;
    }

    try {
      const value = await client.requestProperty('duration');
      const numeric = Number(value);
      deps.setMediaDurationSec(Number.isFinite(numeric) && numeric > 0 ? numeric : null);
    } catch {
      deps.setMediaDurationSec(null);
    }
  };

  const publishDiscordPresence = (): void => {
    const discordPresenceService = deps.getDiscordPresenceService();
    if (!discordPresenceService || deps.isDiscordPresenceEnabled() !== true) {
      return;
    }

    void refreshDiscordPresenceMediaDuration();
    const client = deps.getMpvClient();
    discordPresenceService.publish({
      mediaTitle: deps.getCurrentMediaTitle(),
      mediaPath: deps.getCurrentMediaPath(),
      subtitleText: deps.getCurrentSubtitleText(),
      currentTimeSec: client?.currentTimePos ?? null,
      mediaDurationSec: deps.getMediaDurationSec() ?? deps.getFallbackMediaDurationSec(),
      paused: deps.getPlaybackPaused(),
      connected: Boolean(client?.connected),
      sessionStartedAtMs: deps.getSessionStartedAtMs(),
    });
  };

  return {
    refreshDiscordPresenceMediaDuration,
    publishDiscordPresence,
  };
}

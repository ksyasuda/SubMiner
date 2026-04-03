import { createDiscordPresenceRuntime } from './runtime/discord-presence-runtime';

type DiscordPresenceConfigLike = {
  enabled?: boolean;
};

type DiscordPresenceServiceLike = {
  start: () => Promise<void>;
  stop?: () => Promise<void>;
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

export interface DiscordPresenceLifecycleRuntimeInput {
  getResolvedConfig: () => { discordPresence: DiscordPresenceConfigLike };
  getDiscordPresenceService: () => DiscordPresenceServiceLike | null;
  setDiscordPresenceService: (service: DiscordPresenceServiceLike | null) => void;
  getMpvClient: () => {
    connected?: boolean;
    currentTimePos?: number | null;
    requestProperty: (name: string) => Promise<unknown>;
  } | null;
  getCurrentMediaTitle: () => string | null;
  getCurrentMediaPath: () => string | null;
  getCurrentSubtitleText: () => string;
  getPlaybackPaused: () => boolean | null;
  getFallbackMediaDurationSec: () => number | null;
  createDiscordPresenceService: (config: unknown) => DiscordPresenceServiceLike;
  createDiscordRuntime?: typeof createDiscordPresenceRuntime;
  now?: () => number;
}

export interface DiscordPresenceLifecycleRuntime {
  publishDiscordPresence: () => void;
  initializeDiscordPresenceService: () => Promise<void>;
  stopDiscordPresenceService: () => Promise<void>;
}

export function createDiscordPresenceLifecycleRuntime(
  input: DiscordPresenceLifecycleRuntimeInput,
): DiscordPresenceLifecycleRuntime {
  let discordPresenceMediaDurationSec: number | null = null;
  const discordPresenceSessionStartedAtMs = input.now ? input.now() : Date.now();
  const stopCurrentDiscordPresenceService = async (): Promise<void> => {
    await input.getDiscordPresenceService()?.stop?.();
    input.setDiscordPresenceService(null);
  };

  const discordPresenceRuntime = (input.createDiscordRuntime ?? createDiscordPresenceRuntime)({
    getDiscordPresenceService: () => input.getDiscordPresenceService(),
    isDiscordPresenceEnabled: () => input.getResolvedConfig().discordPresence.enabled === true,
    getMpvClient: () => input.getMpvClient(),
    getCurrentMediaTitle: () => input.getCurrentMediaTitle(),
    getCurrentMediaPath: () => input.getCurrentMediaPath(),
    getCurrentSubtitleText: () => input.getCurrentSubtitleText(),
    getPlaybackPaused: () => input.getPlaybackPaused(),
    getFallbackMediaDurationSec: () => input.getFallbackMediaDurationSec(),
    getSessionStartedAtMs: () => discordPresenceSessionStartedAtMs,
    getMediaDurationSec: () => discordPresenceMediaDurationSec,
    setMediaDurationSec: (next) => {
      discordPresenceMediaDurationSec = next;
    },
  });

  return {
    publishDiscordPresence: () => {
      discordPresenceRuntime.publishDiscordPresence();
    },
    initializeDiscordPresenceService: async () => {
      if (input.getResolvedConfig().discordPresence.enabled !== true) {
        await stopCurrentDiscordPresenceService();
        return;
      }

      await stopCurrentDiscordPresenceService();
      input.setDiscordPresenceService(
        input.createDiscordPresenceService(input.getResolvedConfig().discordPresence),
      );
      await input.getDiscordPresenceService()?.start();
      discordPresenceRuntime.publishDiscordPresence();
    },
    stopDiscordPresenceService: stopCurrentDiscordPresenceService,
  };
}

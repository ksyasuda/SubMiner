import { createDiscordPresenceService } from '../../core/services/discord-presence';
import type { ResolvedConfig } from '../../types';
import { createDiscordRpcClient } from './discord-rpc-client.js';

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

export function createDiscordPresenceRuntimeFromMainState(input: {
  appId: string;
  appState: {
    discordPresenceService: ReturnType<typeof createDiscordPresenceService> | null;
    mpvClient: MpvClientLike | null;
    currentMediaTitle: string | null;
    currentMediaPath: string | null;
    currentSubText: string;
    playbackPaused: boolean | null;
  };
  getResolvedConfig: () => ResolvedConfig;
  getFallbackMediaDurationSec: () => number | null;
  logger: {
    debug: (message: string, meta?: unknown) => void;
  };
}) {
  const sessionStartedAtMs = Date.now();
  let mediaDurationSec: number | null = null;
  const stopCurrentDiscordPresenceService = async (): Promise<void> => {
    await input.appState.discordPresenceService?.stop?.();
    input.appState.discordPresenceService = null;
  };

  const discordPresenceRuntime = createDiscordPresenceRuntime({
    getDiscordPresenceService: () => input.appState.discordPresenceService,
    isDiscordPresenceEnabled: () => input.getResolvedConfig().discordPresence.enabled === true,
    getMpvClient: () => input.appState.mpvClient,
    getCurrentMediaTitle: () => input.appState.currentMediaTitle,
    getCurrentMediaPath: () => input.appState.currentMediaPath,
    getCurrentSubtitleText: () => input.appState.currentSubText,
    getPlaybackPaused: () => input.appState.playbackPaused,
    getFallbackMediaDurationSec: () => input.getFallbackMediaDurationSec(),
    getSessionStartedAtMs: () => sessionStartedAtMs,
    getMediaDurationSec: () => mediaDurationSec,
    setMediaDurationSec: (next) => {
      mediaDurationSec = next;
    },
  });

  const initializeDiscordPresenceService = async (): Promise<void> => {
    if (input.getResolvedConfig().discordPresence.enabled !== true) {
      await stopCurrentDiscordPresenceService();
      return;
    }

    await stopCurrentDiscordPresenceService();
    input.appState.discordPresenceService = createDiscordPresenceService({
      config: input.getResolvedConfig().discordPresence,
      createClient: () => createDiscordRpcClient(input.appId),
      logDebug: (message, meta) => input.logger.debug(message, meta),
    });
    await input.appState.discordPresenceService.start();
    discordPresenceRuntime.publishDiscordPresence();
  };

  return {
    discordPresenceRuntime,
    initializeDiscordPresenceService,
  };
}

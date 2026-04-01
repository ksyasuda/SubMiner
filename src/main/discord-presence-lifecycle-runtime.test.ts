import assert from 'node:assert/strict';
import test from 'node:test';

import { createDiscordPresenceLifecycleRuntime } from './discord-presence-lifecycle-runtime';

test('discord presence lifecycle runtime starts service and publishes presence when enabled', async () => {
  const calls: string[] = [];
  let service: { start: () => Promise<void>; stop: () => Promise<void> } | null = null;

  const runtime = createDiscordPresenceLifecycleRuntime({
    getResolvedConfig: () => ({ discordPresence: { enabled: true } }),
    getDiscordPresenceService: () => service as never,
    setDiscordPresenceService: (next) => {
      service = next as typeof service;
    },
    getMpvClient: () => null,
    getCurrentMediaTitle: () => 'Demo',
    getCurrentMediaPath: () => '/tmp/demo.mkv',
    getCurrentSubtitleText: () => 'subtitle',
    getPlaybackPaused: () => false,
    getFallbackMediaDurationSec: () => 12,
    createDiscordPresenceService: () => ({
      start: async () => {
        calls.push('start');
      },
      stop: async () => {
        calls.push('stop');
      },
      publish: () => {
        calls.push('publish');
      },
    }),
    createDiscordRuntime: (input) => ({
      refreshDiscordPresenceMediaDuration: async () => {},
      publishDiscordPresence: () => {
        calls.push(input.getCurrentMediaTitle() ?? 'unknown');
        input.getDiscordPresenceService()?.publish({
          mediaTitle: input.getCurrentMediaTitle(),
          mediaPath: input.getCurrentMediaPath(),
          subtitleText: input.getCurrentSubtitleText(),
          currentTimeSec: null,
          mediaDurationSec: input.getFallbackMediaDurationSec(),
          paused: input.getPlaybackPaused(),
          connected: false,
          sessionStartedAtMs: input.getSessionStartedAtMs(),
        });
      },
    }),
    now: () => 123,
  });

  await runtime.initializeDiscordPresenceService();

  assert.deepEqual(calls, ['start', 'Demo', 'publish']);
});

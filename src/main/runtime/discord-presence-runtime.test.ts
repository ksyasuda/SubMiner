import assert from 'node:assert/strict';
import test from 'node:test';
import { createDiscordPresenceRuntime } from './discord-presence-runtime';

test('discord presence runtime refreshes duration and publishes the current snapshot', async () => {
  const snapshots: Array<Record<string, unknown>> = [];
  let mediaDurationSec: number | null = null;

  const runtime = createDiscordPresenceRuntime({
    getDiscordPresenceService: () => ({
      publish: (snapshot: Record<string, unknown>) => {
        snapshots.push(snapshot);
      },
    }),
    isDiscordPresenceEnabled: () => true,
    getMpvClient: () =>
      ({
        connected: true,
        currentTimePos: 12,
        requestProperty: async (name: string) => {
          assert.equal(name, 'duration');
          return 42;
        },
      }) as never,
    getCurrentMediaTitle: () => 'Episode 1',
    getCurrentMediaPath: () => '/media/episode-1.mkv',
    getCurrentSubtitleText: () => '字幕',
    getPlaybackPaused: () => false,
    getFallbackMediaDurationSec: () => 90,
    getSessionStartedAtMs: () => 1_000,
    getMediaDurationSec: () => mediaDurationSec,
    setMediaDurationSec: (next) => {
      mediaDurationSec = next;
    },
  });

  await runtime.refreshDiscordPresenceMediaDuration();
  runtime.publishDiscordPresence();

  assert.equal(mediaDurationSec, 42);
  assert.deepEqual(snapshots, [
    {
      mediaTitle: 'Episode 1',
      mediaPath: '/media/episode-1.mkv',
      subtitleText: '字幕',
      currentTimeSec: 12,
      mediaDurationSec: 42,
      paused: false,
      connected: true,
      sessionStartedAtMs: 1_000,
    },
  ]);
});

test('discord presence runtime skips publish when disabled or service missing', () => {
  let published = false;
  const runtime = createDiscordPresenceRuntime({
    getDiscordPresenceService: () => null,
    isDiscordPresenceEnabled: () => false,
    getMpvClient: () => null,
    getCurrentMediaTitle: () => null,
    getCurrentMediaPath: () => null,
    getCurrentSubtitleText: () => '',
    getPlaybackPaused: () => null,
    getFallbackMediaDurationSec: () => null,
    getSessionStartedAtMs: () => 0,
    getMediaDurationSec: () => null,
    setMediaDurationSec: () => {
      published = true;
    },
  });

  runtime.publishDiscordPresence();

  assert.equal(published, false);
});

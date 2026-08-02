import test from 'node:test';
import assert from 'node:assert/strict';
import type { AnimeBridgeClient, BridgeSource } from '../../anime-bridge/bridge-client';
import type { AnimeBrowserPlayRequest } from '../../types/anime-browser';
import { createAnimeBrowserPlayback } from './anime-browser-playback';

const request: AnimeBrowserPlayRequest = {
  sourceId: 'source',
  animeUrl: '/anime',
  animeTitle: 'Anime',
  episodeUrl: '/episode-1',
  episodeName: 'Episode 1',
  episodeNumber: 1,
};

test('playback returns the established error when an extension has no playable stream', async () => {
  let mpvConnections = 0;
  const client = { getVideoList: async () => [] } as unknown as AnimeBridgeClient;
  const playback = createAnimeBrowserPlayback({
    deps: {
      sendMpvCommand: () => undefined,
      ensureMpvConnected: async () => {
        mpvConnections += 1;
        return true;
      },
      log: () => undefined,
    },
    bridge: async () => ({ client, baseUrl: 'http://127.0.0.1:1234' }),
    sourceFor: async () => ({}) as BridgeSource,
    stripProxy: () => null,
  });

  assert.deepEqual(await playback.playEpisode(request), {
    ok: false,
    error: 'That source returned no playable video.',
    quality: null,
  });
  assert.equal(mpvConnections, 0);
  await playback.dispose();
});

test('a superseded episode stops instead of driving mpv behind the newer one', async () => {
  let releaseFirst = (): void => undefined;
  const firstResolved = new Promise<void>((resolve) => {
    releaseFirst = resolve;
  });
  let calls = 0;
  const client = {
    async getVideoList() {
      calls += 1;
      if (calls === 1) await firstResolved;
      return [{ videoUrl: `http://stream/${calls}.m3u8`, quality: '1080p' }];
    },
  } as unknown as AnimeBridgeClient;

  const commands: Array<Array<string | number>> = [];
  let overlays = 0;
  const playback = createAnimeBrowserPlayback({
    deps: {
      sendMpvCommand: (command) => void commands.push(command),
      ensureMpvConnected: async () => true,
      log: () => undefined,
      showVisibleOverlay: () => {
        overlays += 1;
      },
    },
    bridge: async () => ({ client, baseUrl: 'http://127.0.0.1:1234' }),
    sourceFor: async () => ({}) as BridgeSource,
    stripProxy: () => null,
  });

  const stale = playback.playEpisode(request);
  const fresh = await playback.playEpisode({ ...request, episodeUrl: '/episode-2' });
  const commandsAfterFresh = commands.length;
  releaseFirst();

  assert.equal(fresh.ok, true);
  assert.equal(overlays, 1);
  assert.deepEqual(await stale, {
    ok: false,
    error: 'A newer episode replaced this playback.',
    quality: null,
  });
  assert.equal(commands.length, commandsAfterFresh);
  assert.equal(overlays, 1);
  await playback.dispose();
});

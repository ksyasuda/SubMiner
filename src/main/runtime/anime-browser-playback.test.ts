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

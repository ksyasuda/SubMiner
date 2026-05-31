import assert from 'node:assert/strict';
import test from 'node:test';

import { resolveFreshPlaybackPaused } from './playback-paused-state';

test('resolveFreshPlaybackPaused prefers the live mpv pause property over cached state', async () => {
  const paused = await resolveFreshPlaybackPaused({
    getCachedPlaybackPaused: () => false,
    getMpvClient: () => ({
      connected: true,
      requestProperty: async (name: string) => (name === 'pause' ? true : null),
    }),
  });

  assert.equal(paused, true);
});

test('resolveFreshPlaybackPaused normalizes mpv pause property strings and numbers', async () => {
  const values: Array<[unknown, boolean]> = [
    ['yes', true],
    ['no', false],
    ['0', false],
    [1, true],
    [0, false],
  ];

  for (const [value, expected] of values) {
    const paused = await resolveFreshPlaybackPaused({
      getCachedPlaybackPaused: () => null,
      getMpvClient: () => ({
        connected: true,
        requestProperty: async () => value,
      }),
    });

    assert.equal(paused, expected);
  }
});

test('resolveFreshPlaybackPaused falls back to cached state when mpv is unavailable', async () => {
  assert.equal(
    await resolveFreshPlaybackPaused({
      getCachedPlaybackPaused: () => true,
      getMpvClient: () => null,
    }),
    true,
  );
  assert.equal(
    await resolveFreshPlaybackPaused({
      getCachedPlaybackPaused: () => false,
      getMpvClient: () => ({
        connected: true,
        requestProperty: async () => {
          throw new Error('socket closed');
        },
      }),
    }),
    false,
  );
});

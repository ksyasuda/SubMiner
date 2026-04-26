import assert from 'node:assert/strict';
import test from 'node:test';

import { searchAniListMediaCandidates } from './fetch';

test('searchAniListMediaCandidates trims fallback candidate titles', async () => {
  const previousFetchDescriptor = Object.getOwnPropertyDescriptor(globalThis, 'fetch');
  Object.defineProperty(globalThis, 'fetch', {
    configurable: true,
    writable: true,
    value: async () =>
      new Response(
        JSON.stringify({
          data: {
            Page: {
              media: [{ id: 21355, episodes: 25, title: {} }],
            },
          },
        }),
      ),
  });

  try {
    const candidates = await searchAniListMediaCandidates('  Re:ZERO  ');

    assert.equal(candidates[0]?.title, 'Re:ZERO');
  } finally {
    if (previousFetchDescriptor) {
      Object.defineProperty(globalThis, 'fetch', previousFetchDescriptor);
    } else {
      Reflect.deleteProperty(globalThis, 'fetch');
    }
  }
});

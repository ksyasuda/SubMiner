import test from 'node:test';
import assert from 'node:assert/strict';

import { guessAnilistMediaInfo, updateAnilistPostWatchProgress } from './anilist-updater';

function createJsonResponse(payload: unknown): Response {
  return new Response(JSON.stringify(payload), {
    status: 200,
    headers: { 'content-type': 'application/json' },
  });
}

test('guessAnilistMediaInfo uses guessit output when available', async () => {
  const result = await guessAnilistMediaInfo('/tmp/demo.mkv', null, {
    runGuessit: async () => JSON.stringify({ title: 'Guessit Title', episode: 7 }),
  });
  assert.deepEqual(result, {
    title: 'Guessit Title',
    episode: 7,
    source: 'guessit',
  });
});

test('guessAnilistMediaInfo falls back to parser when guessit fails', async () => {
  const result = await guessAnilistMediaInfo('/tmp/My Anime S01E03.mkv', null, {
    runGuessit: async () => {
      throw new Error('guessit not found');
    },
  });
  assert.deepEqual(result, {
    title: 'My Anime',
    episode: 3,
    source: 'fallback',
  });
});

test('guessAnilistMediaInfo uses basename for guessit input', async () => {
  const mediaPath =
    '/truenas/jellyfin/anime/Rascal-Does-not-Dream-of-Bunny-Girl-Senapi/Season-1/Rascal Does Not Dream of Bunny Girl Senpai (2018) - S01E01 - 001 - My Senpai Is a Bunny Girl [Bluray-1080p][10bit][x265][Opus 2.0][JA]-Subs.mkv';
  const seenTargets: string[] = [];
  const result = await guessAnilistMediaInfo(mediaPath, null, {
    runGuessit: async (target) => {
      seenTargets.push(target);
      return JSON.stringify({
        title: 'Rascal Does Not Dream of Bunny Girl Senpai',
        episode: 1,
      });
    },
  });
  assert.deepEqual(seenTargets, [
    'Rascal Does Not Dream of Bunny Girl Senpai (2018) - S01E01 - 001 - My Senpai Is a Bunny Girl [Bluray-1080p][10bit][x265][Opus 2.0][JA]-Subs.mkv',
  ]);
  assert.deepEqual(result, {
    title: 'Rascal Does Not Dream of Bunny Girl Senpai',
    episode: 1,
    source: 'guessit',
  });
});

test('guessAnilistMediaInfo joins multi-part guessit titles', async () => {
  const result = await guessAnilistMediaInfo('/tmp/demo.mkv', null, {
    runGuessit: async () =>
      JSON.stringify({
        title: ['Rascal', 'Does-not-Dream-of-Bunny-Girl-Senpai'],
        episode: 1,
      }),
  });
  assert.deepEqual(result, {
    title: 'Rascal Does not Dream of Bunny Girl Senpai',
    episode: 1,
    source: 'guessit',
  });
});

test('updateAnilistPostWatchProgress updates progress when behind', async () => {
  const originalFetch = globalThis.fetch;
  let call = 0;
  globalThis.fetch = (async () => {
    call += 1;
    if (call === 1) {
      return createJsonResponse({
        data: {
          Page: {
            media: [
              {
                id: 11,
                episodes: 24,
                title: { english: 'Demo Show', romaji: 'Demo Show' },
              },
            ],
          },
        },
      });
    }
    if (call === 2) {
      return createJsonResponse({
        data: {
          Media: {
            id: 11,
            mediaListEntry: { progress: 2, status: 'CURRENT' },
          },
        },
      });
    }
    return createJsonResponse({
      data: { SaveMediaListEntry: { progress: 3, status: 'CURRENT' } },
    });
  }) as typeof fetch;

  try {
    const result = await updateAnilistPostWatchProgress('token', 'Demo Show', 3);
    assert.equal(result.status, 'updated');
    assert.match(result.message, /episode 3/i);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('updateAnilistPostWatchProgress skips when progress already reached', async () => {
  const originalFetch = globalThis.fetch;
  let call = 0;
  globalThis.fetch = (async () => {
    call += 1;
    if (call === 1) {
      return createJsonResponse({
        data: {
          Page: {
            media: [{ id: 22, episodes: 12, title: { english: 'Skip Show' } }],
          },
        },
      });
    }
    return createJsonResponse({
      data: {
        Media: { id: 22, mediaListEntry: { progress: 12, status: 'CURRENT' } },
      },
    });
  }) as typeof fetch;

  try {
    const result = await updateAnilistPostWatchProgress('token', 'Skip Show', 10);
    assert.equal(result.status, 'skipped');
    assert.match(result.message, /already at episode/i);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('updateAnilistPostWatchProgress returns error when search fails', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async () =>
    createJsonResponse({
      errors: [{ message: 'bad request' }],
    })) as typeof fetch;

  try {
    const result = await updateAnilistPostWatchProgress('token', 'Bad', 1);
    assert.equal(result.status, 'error');
    assert.match(result.message, /search failed/i);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

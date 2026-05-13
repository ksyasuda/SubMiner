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
    season: null,
    episode: 7,
    source: 'guessit',
  });
});

test('guessAnilistMediaInfo fills missing guessit episode from filename parser', async () => {
  const result = await guessAnilistMediaInfo('/tmp/Guessit Title S01E09.mkv', null, {
    runGuessit: async () => JSON.stringify({ title: 'Guessit Title' }),
  });
  assert.deepEqual(result, {
    title: 'Guessit Title',
    season: 1,
    episode: 9,
    source: 'guessit',
  });
});

test('guessAnilistMediaInfo ignores low-confidence parser details when guessit omits them', async () => {
  const result = await guessAnilistMediaInfo('/tmp/Season 2/Guessit Title.mkv', null, {
    runGuessit: async () => JSON.stringify({ title: 'Guessit Title' }),
  });
  assert.deepEqual(result, {
    title: 'Guessit Title',
    season: null,
    episode: null,
    source: 'guessit',
  });
});

test('guessAnilistMediaInfo parses Little Witch Academia release filename', async () => {
  const filename =
    '/tmp/Little Witch Academia (2017) - S01E02 - 002 - Papiliodia [Bluray-1080p][10bit][h265][AC3 2.0][JA].mkv';
  const result = await guessAnilistMediaInfo(filename, null, {
    runGuessit: async () => JSON.stringify({ title: 'Little Witch Academia' }),
  });
  assert.deepEqual(result, {
    title: 'Little Witch Academia',
    season: 1,
    episode: 2,
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
    season: 1,
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
    season: 1,
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
    season: null,
    episode: 1,
    source: 'guessit',
  });
});

test('guessAnilistMediaInfo preserves useful guessit alternative title for ambiguous Re ZERO filenames', async () => {
  const result = await guessAnilistMediaInfo(
    '/tmp/Re - ZERO, Starting Life in Another World (2016) - S01E01 - - The End of the Beginning and the Beginning of the End [v2 Bluray-1080p Proper][10bit][x265][FLAC 2.0][EN+JA]-SCY.mkv',
    null,
    {
      runGuessit: async () =>
        JSON.stringify({
          title: 'Re',
          alternative_title: 'ZERO, Starting Life in Another World',
          year: 2016,
          season: 1,
          episode: 1,
        }),
    },
  );

  assert.deepEqual(result, {
    title: 'Re ZERO, Starting Life in Another World',
    alternativeTitle: 'ZERO, Starting Life in Another World',
    year: 2016,
    season: 1,
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

test('updateAnilistPostWatchProgress uses the configured AniList rate limiter', async () => {
  const originalFetch = globalThis.fetch;
  let call = 0;
  let acquireCalls = 0;
  let recordCalls = 0;
  globalThis.fetch = (async () => {
    call += 1;
    if (call === 1) {
      return createJsonResponse({
        data: {
          Page: {
            media: [{ id: 11, episodes: 24, title: { english: 'Demo Show' } }],
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
    const result = await updateAnilistPostWatchProgress('token', 'Demo Show', 3, {
      rateLimiter: {
        acquire: async () => {
          acquireCalls += 1;
        },
        recordResponse: () => {
          recordCalls += 1;
        },
      },
    });

    assert.equal(result.status, 'updated');
    assert.equal(acquireCalls, 3);
    assert.equal(recordCalls, 3);
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

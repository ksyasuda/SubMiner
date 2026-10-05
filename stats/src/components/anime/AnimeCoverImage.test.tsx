import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { AnimeHeader } from './AnimeHeader';

test('AnimeHeader uses the linked AniList id to avoid stale cached cover art', () => {
  const markup = renderToStaticMarkup(
    <AnimeHeader
      detail={{
        mediaKind: 'anime',
        animeId: 42,
        canonicalTitle: 'Test Anime',
        anilistId: 21699,
        tmdbId: null,
        tmdbType: null,
        titleRomaji: null,
        titleEnglish: null,
        titleNative: null,
        description: null,
        totalSessions: 0,
        totalActiveMs: 0,
        totalCards: 0,
        totalTokensSeen: 0,
        totalLinesSeen: 0,
        totalLookupCount: 0,
        totalLookupHits: 0,
        totalYomitanLookupCount: 0,
        episodeCount: 1,
        lastWatchedMs: 0,
      }}
      anilistEntries={[]}
    />,
  );

  assert.match(markup, /\/api\/stats\/anime\/42\/cover\?coverRetry=21699/);
});

import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { AnimeCoverImage } from './AnimeCoverImage';
import { AnimeHeader } from './AnimeHeader';

test('AnimeCoverImage includes manual relink cover retry tokens', () => {
  const markup = renderToStaticMarkup(
    <AnimeCoverImage animeId={42} title="Test Anime" coverRetryToken={7} />,
  );

  assert.match(markup, /\/api\/stats\/anime\/42\/cover\?coverRetry=7/);
});

test('AnimeHeader uses the linked AniList id to avoid stale cached cover art', () => {
  const markup = renderToStaticMarkup(
    <AnimeHeader
      detail={{
        animeId: 42,
        canonicalTitle: 'Test Anime',
        anilistId: 21699,
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

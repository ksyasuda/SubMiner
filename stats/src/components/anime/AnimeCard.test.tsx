import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { AnimeCard } from './AnimeCard';

test('AnimeCard includes linked AniList id in cover URLs to avoid stale library covers', () => {
  const markup = renderToStaticMarkup(
    <AnimeCard
      anime={{
        animeId: 42,
        canonicalTitle: 'Test Anime',
        anilistId: 21699,
        totalSessions: 1,
        totalActiveMs: 600_000,
        totalCards: 0,
        totalTokensSeen: 100,
        episodeCount: 1,
        episodesTotal: 10,
        lastWatchedMs: 1_000,
      }}
      onClick={() => {}}
    />,
  );

  assert.match(markup, /\/api\/stats\/anime\/42\/cover\?coverRetry=21699/);
});

import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { MediaHeader } from '../components/library/MediaHeader';
import { EpisodeList } from '../components/anime/EpisodeList';
import { AnimeOverviewStats } from '../components/anime/AnimeOverviewStats';
import { SessionRow } from '../components/sessions/SessionRow';
import { EventType, type SessionEvent } from '../types/stats';
import { buildLookupRateDisplay, getYomitanLookupEvents } from './yomitan-lookup';

test('buildLookupRateDisplay formats lookups per 100 tokens in short and long forms', () => {
  assert.deepEqual(buildLookupRateDisplay(23, 1000), {
    shortValue: '2.3 / 100 tokens',
    longValue: '2.3 lookups per 100 tokens',
  });
  assert.equal(buildLookupRateDisplay(0, 0), null);
});

test('getYomitanLookupEvents keeps only Yomitan lookup events', () => {
  const events: SessionEvent[] = [
    { eventType: EventType.LOOKUP, tsMs: 1, payload: null },
    { eventType: EventType.YOMITAN_LOOKUP, tsMs: 2, payload: null },
    { eventType: EventType.CARD_MINED, tsMs: 3, payload: null },
  ];

  assert.deepEqual(
    getYomitanLookupEvents(events).map((event) => event.tsMs),
    [2],
  );
});

test('MediaHeader renders Yomitan lookup count and lookup rate copy', () => {
  const markup = renderToStaticMarkup(
    <MediaHeader
      detail={{
        videoId: 7,
        canonicalTitle: 'Episode 7',
        totalSessions: 4,
        totalActiveMs: 90_000,
        totalCards: 12,
        totalTokensSeen: 1000,
        totalLinesSeen: 120,
        totalLookupCount: 30,
        totalLookupHits: 21,
        totalYomitanLookupCount: 23,
      }}
    />,
  );

  assert.match(markup, /23/);
  assert.match(markup, /2\.3 \/ 100 tokens/);
  assert.match(markup, /2\.3 lookups per 100 tokens/);
});

test('MediaHeader distinguishes token occurrences from known unique words', () => {
  const markup = renderToStaticMarkup(
    <MediaHeader
      detail={{
        videoId: 7,
        canonicalTitle: 'Episode 7',
        totalSessions: 4,
        totalActiveMs: 90_000,
        totalCards: 12,
        totalTokensSeen: 30,
        totalLinesSeen: 120,
        totalLookupCount: 30,
        totalLookupHits: 21,
        totalYomitanLookupCount: 0,
      }}
      initialKnownWordsSummary={{
        knownWordCount: 17,
        totalUniqueWords: 34,
      }}
    />,
  );

  assert.match(markup, /token occurrences/);
  assert.match(markup, /known unique words \(50%\)/);
  assert.match(markup, /17 \/ 34/);
});

test('EpisodeList renders per-episode Yomitan lookup rate', () => {
  const markup = renderToStaticMarkup(
    <EpisodeList
      episodes={[
        {
          videoId: 9,
          episode: 9,
          season: 1,
          durationMs: 1,
          watched: 0,
          canonicalTitle: 'Episode 9',
          totalSessions: 1,
          totalActiveMs: 1,
          totalCards: 1,
          totalTokensSeen: 350,
          totalYomitanLookupCount: 7,
          lastWatchedMs: 0,
        },
      ]}
    />,
  );

  assert.match(markup, /Lookup Rate/);
  assert.match(markup, /2\.0 \/ 100 tokens/);
});

test('AnimeOverviewStats renders aggregate Yomitan lookup metrics', () => {
  const markup = renderToStaticMarkup(
    <AnimeOverviewStats
      detail={{
        animeId: 1,
        canonicalTitle: 'Anime',
        anilistId: null,
        titleRomaji: null,
        titleEnglish: null,
        titleNative: null,
        description: null,
        totalSessions: 5,
        totalActiveMs: 100_000,
        totalCards: 8,
        totalTokensSeen: 800,
        totalLinesSeen: 100,
        totalLookupCount: 50,
        totalLookupHits: 30,
        totalYomitanLookupCount: 16,
        episodeCount: 3,
        lastWatchedMs: 0,
      }}
      avgSessionMs={20_000}
      knownWordsSummary={null}
    />,
  );

  assert.match(markup, /Lookups/);
  assert.match(markup, /16/);
  assert.match(markup, /2\.0 \/ 100 tokens/);
  assert.match(markup, /Yomitan lookups per 100 tokens seen/);
});

test('SessionRow prefers token-based word count when available', () => {
  const markup = renderToStaticMarkup(
    <SessionRow
      session={{
        sessionId: 7,
        canonicalTitle: 'Episode 7',
        videoId: 7,
        animeId: null,
        animeTitle: null,
        startedAtMs: 0,
        endedAtMs: null,
        totalWatchedMs: 0,
        activeWatchedMs: 0,
        linesSeen: 12,
        tokensSeen: 42,
        cardsMined: 0,
        lookupCount: 0,
        lookupHits: 0,
        yomitanLookupCount: 0,
        knownWordsSeen: 0,
        knownWordRate: 0,
      }}
      isExpanded={false}
      detailsId="session-7"
      onToggle={() => {}}
      onDelete={() => {}}
    />,
  );

  assert.match(markup, />42</);
  assert.doesNotMatch(markup, />12</);
});

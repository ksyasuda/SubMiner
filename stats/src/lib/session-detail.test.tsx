import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { SessionDetail, getKnownPctAxisMax } from '../components/sessions/SessionDetail';
import { buildSessionChartEvents } from './session-events';
import { EventType } from '../types/stats';

test('SessionDetail omits the misleading new words metric', () => {
  const markup = renderToStaticMarkup(
    <SessionDetail
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
        tokensSeen: 24,
        cardsMined: 0,
        lookupCount: 0,
        lookupHits: 0,
        yomitanLookupCount: 0,
        knownWordsSeen: 0,
        knownWordRate: 0,
      }}
    />,
  );

  assert.match(markup, /No token data/);
  assert.doesNotMatch(markup, /New words/);
});

test('buildSessionChartEvents keeps only chart-relevant events and pairs pause ranges', () => {
  const chartEvents = buildSessionChartEvents([
    { eventType: EventType.SUBTITLE_LINE, tsMs: 1_000, payload: '{"line":"ignored"}' },
    { eventType: EventType.PAUSE_START, tsMs: 2_000, payload: null },
    { eventType: EventType.SEEK_FORWARD, tsMs: 3_000, payload: null },
    { eventType: EventType.PAUSE_END, tsMs: 4_000, payload: null },
    { eventType: EventType.CARD_MINED, tsMs: 5_000, payload: '{"cardsMined":1}' },
    { eventType: EventType.YOMITAN_LOOKUP, tsMs: 6_000, payload: null },
    { eventType: EventType.SEEK_BACKWARD, tsMs: 7_000, payload: null },
    { eventType: EventType.LOOKUP, tsMs: 8_000, payload: '{"hit":true}' },
  ]);

  assert.deepEqual(
    chartEvents.seekEvents.map((event) => event.eventType),
    [EventType.SEEK_FORWARD, EventType.SEEK_BACKWARD],
  );
  assert.deepEqual(
    chartEvents.cardEvents.map((event) => event.tsMs),
    [5_000],
  );
  assert.deepEqual(
    chartEvents.yomitanLookupEvents.map((event) => event.tsMs),
    [6_000],
  );
  assert.deepEqual(chartEvents.pauseRegions, [{ startMs: 2_000, endMs: 4_000 }]);
});

test('getKnownPctAxisMax adds headroom above the highest known percentage', () => {
  assert.equal(getKnownPctAxisMax([22.4, 31.2, 29.8]), 40);
});

test('getKnownPctAxisMax caps the chart top at 100%', () => {
  assert.equal(getKnownPctAxisMax([97.1, 98.6]), 100);
});

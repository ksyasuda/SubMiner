import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { TrackingSnapshot } from './TrackingSnapshot';
import type { OverviewSummary } from '../../lib/dashboard-data';

const summary: OverviewSummary = {
  todayActiveMs: 0,
  todayCards: 0,
  streakDays: 0,
  allTimeMinutes: 120,
  totalTrackedCards: 9,
  episodesToday: 0,
  activeAnimeCount: 0,
  totalEpisodesWatched: 5,
  totalAnimeCompleted: 1,
  averageSessionMinutes: 40,
  activeDays: 12,
  totalSessions: 15,
  lookupRate: {
    shortValue: '2.3 / 100 words',
    longValue: '2.3 lookups per 100 words',
  },
  todayTokens: 0,
  newWordsToday: 0,
  newWordsThisWeek: 0,
  recentWatchTime: [],
};

test('TrackingSnapshot renders Yomitan lookup rate copy on the homepage card', () => {
  const markup = renderToStaticMarkup(
    <TrackingSnapshot summary={summary} knownWordsSummary={null} />,
  );

  assert.match(markup, /Lookup Rate/);
  assert.match(markup, /2\.3 \/ 100 words/);
  assert.match(markup, /Lifetime Yomitan lookups normalized by total words seen/);
});

test('TrackingSnapshot labels new words as unique headwords', () => {
  const markup = renderToStaticMarkup(
    <TrackingSnapshot summary={summary} knownWordsSummary={null} />,
  );

  assert.match(markup, /Unique headwords seen for the first time today/);
  assert.match(markup, /Unique headwords seen for the first time this week/);
});

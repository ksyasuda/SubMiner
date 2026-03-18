import assert from 'node:assert/strict';
import test from 'node:test';

import type {
  DailyRollup,
  OverviewData,
  SessionSummary,
  StreakCalendarDay,
  VocabularyEntry,
} from '../types/stats';
import {
  buildOverviewSummary,
  buildStreakCalendar,
  buildTrendDashboard,
  buildVocabularySummary,
} from './dashboard-data';

test('buildOverviewSummary aggregates tracked totals and recent windows', () => {
  const now = Date.UTC(2026, 2, 13, 12);
  const today = Math.floor(now / 86_400_000);
  const sessions: SessionSummary[] = [
    {
      sessionId: 1,
      canonicalTitle: 'A',
      videoId: 1,
      animeId: null,
      animeTitle: null,
      startedAtMs: now - 3_600_000,
      endedAtMs: now - 1_800_000,
      totalWatchedMs: 3_600_000,
      activeWatchedMs: 3_000_000,
      linesSeen: 20,
      tokensSeen: 80,
      cardsMined: 2,
      lookupCount: 10,
      lookupHits: 8,
      yomitanLookupCount: 0,
    },
  ];
  const rollups: DailyRollup[] = [
    {
      rollupDayOrMonth: today,
      videoId: 1,
      totalSessions: 1,
      totalActiveMin: 50,
      totalLinesSeen: 20,
      totalTokensSeen: 80,
      totalCards: 2,
      cardsPerHour: 2.4,
      tokensPerMin: 2,
      lookupHitRate: 0.8,
    },
  ];
  const overview: OverviewData = {
    sessions,
    rollups,
    hints: {
      totalSessions: 15,
      activeSessions: 0,
      episodesToday: 2,
      activeAnimeCount: 3,
      totalEpisodesWatched: 5,
      totalAnimeCompleted: 1,
      totalActiveMin: 50,
      activeDays: 2,
      totalCards: 9,
      totalLookupCount: 100,
      totalLookupHits: 80,
      newWordsToday: 5,
      newWordsThisWeek: 20,
    },
  };

  const summary = buildOverviewSummary(overview, now);
  assert.equal(summary.todayCards, 2);
  assert.equal(summary.totalTrackedCards, 9);
  assert.equal(summary.episodesToday, 2);
  assert.equal(summary.activeAnimeCount, 3);
  assert.equal(summary.averageSessionMinutes, 50);
  assert.equal(summary.allTimeMinutes, 50);
  assert.equal(summary.activeDays, 2);
  assert.equal(summary.totalSessions, 15);
  assert.equal(summary.lookupRate, 80);
});

test('buildOverviewSummary prefers lifetime totals from hints when provided', () => {
  const now = Date.UTC(2026, 2, 13, 12);
  const today = Math.floor(now / 86_400_000);
  const overview: OverviewData = {
    sessions: [
      {
        sessionId: 2,
        canonicalTitle: 'B',
        videoId: 2,
        animeId: null,
        animeTitle: null,
        startedAtMs: now - 60_000,
        endedAtMs: now,
        totalWatchedMs: 60_000,
        activeWatchedMs: 60_000,
        linesSeen: 10,
        tokensSeen: 10,
        cardsMined: 10,
        lookupCount: 1,
        lookupHits: 1,
        yomitanLookupCount: 0,
      },
    ],
    rollups: [
      {
        rollupDayOrMonth: today,
        videoId: 2,
        totalSessions: 1,
        totalActiveMin: 1,
        totalLinesSeen: 10,
        totalTokensSeen: 10,
        totalCards: 10,
        cardsPerHour: 600,
        tokensPerMin: 10,
        lookupHitRate: 1,
      },
    ],
    hints: {
      totalSessions: 50,
      activeSessions: 0,
      episodesToday: 0,
      activeAnimeCount: 0,
      totalEpisodesWatched: 0,
      totalAnimeCompleted: 0,
      totalActiveMin: 120,
      activeDays: 40,
      totalCards: 5,
      totalLookupCount: 0,
      totalLookupHits: 0,
      newWordsToday: 0,
      newWordsThisWeek: 0,
    },
  };

  const summary = buildOverviewSummary(overview, now);
  assert.equal(summary.totalTrackedCards, 5);
  assert.equal(summary.allTimeMinutes, 120);
  assert.equal(summary.activeDays, 40);
});

test('buildVocabularySummary treats firstSeen timestamps as seconds', () => {
  const now = Date.UTC(2026, 2, 13, 12);
  const nowSec = now / 1000;
  const words: VocabularyEntry[] = [
    {
      wordId: 1,
      headword: '猫',
      word: '猫',
      reading: 'ねこ',
      partOfSpeech: null,
      pos1: null,
      pos2: null,
      pos3: null,
      frequency: 4,
      frequencyRank: null,
      animeCount: 1,
      firstSeen: nowSec - 2 * 86_400,
      lastSeen: nowSec - 1,
    },
  ];

  const summary = buildVocabularySummary(words, [], now);
  assert.equal(summary.newThisWeek, 1);
});

test('buildTrendDashboard derives dense chart series', () => {
  const now = Date.UTC(2026, 2, 13, 12);
  const today = Math.floor(now / 86_400_000);
  const rollups: DailyRollup[] = [
    {
      rollupDayOrMonth: today - 1,
      videoId: 1,
      totalSessions: 2,
      totalActiveMin: 60,
      totalLinesSeen: 30,
      totalTokensSeen: 100,
      totalCards: 3,
      cardsPerHour: 3,
      tokensPerMin: 2,
      lookupHitRate: 0.5,
    },
    {
      rollupDayOrMonth: today,
      videoId: 1,
      totalSessions: 1,
      totalActiveMin: 30,
      totalLinesSeen: 10,
      totalTokensSeen: 30,
      totalCards: 1,
      cardsPerHour: 2,
      tokensPerMin: 1.33,
      lookupHitRate: 0.75,
    },
  ];

  const dashboard = buildTrendDashboard(rollups);
  assert.equal(dashboard.watchTime.length, 2);
  assert.equal(dashboard.words[1]?.value, 30);
  assert.equal(dashboard.sessions[0]?.value, 2);
});

test('buildStreakCalendar converts epoch days to YYYY-MM-DD dates', () => {
  const days: StreakCalendarDay[] = [
    { epochDay: 20525, totalActiveMin: 45 },
    { epochDay: 20526, totalActiveMin: 0 },
    { epochDay: 20527, totalActiveMin: 30 },
  ];

  const points = buildStreakCalendar(days);
  assert.equal(points.length, 3);
  assert.match(points[0]!.date, /^\d{4}-\d{2}-\d{2}$/);
  assert.equal(points[0]!.value, 45);
  assert.equal(points[1]!.value, 0);
  assert.equal(points[2]!.value, 30);
});

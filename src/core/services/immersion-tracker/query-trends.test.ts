import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { Database, type DatabaseSync } from './sqlite';
import { getOrCreateAnimeRecord, getOrCreateVideoRecord, linkVideoToAnimeRecord } from './storage';
import { ensureSchema } from './storage';
import { SOURCE_TYPE_LOCAL } from './types';
import { getTrendsDashboard } from './query-trends';
import { getSessionSummaries } from './query-sessions';
import { getVocabularyChartData, getVocabularySummary } from './query-lexical';
import { markLexicalDailyRollupsReady } from './lexical-rollups';

/** Noon, local time, on a calendar day, so the local epoch day holds in every time zone. */
function localNoonMs(epochDay: number): number {
  const date = new Date(epochDay * 86_400_000);
  return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(), 12).getTime();
}

// Mocked "now" is noon on 2026-03-01, a Sunday.
const TODAY = Date.UTC(2026, 2, 1) / 86_400_000;
const NOW_MS = String(localNoonMs(TODAY));

function withDashboardDb(run: (db: DatabaseSync) => void): void {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-trends-'));
  const db = new Database(path.join(dir, 'immersion.sqlite'));
  const previousNow = globalThis.__subminerTestNowMs;
  globalThis.__subminerTestNowMs = NOW_MS;
  try {
    ensureSchema(db);
    run(db);
  } finally {
    globalThis.__subminerTestNowMs = previousNow;
    db.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

function addVideo(db: DatabaseSync, title: string, episode: number, watched = false): number {
  const key = `${title}-${episode}`;
  const videoId = getOrCreateVideoRecord(db, `local:/tmp/${key}.mkv`, {
    canonicalTitle: key,
    sourcePath: `/tmp/${key}.mkv`,
    sourceUrl: null,
    sourceType: SOURCE_TYPE_LOCAL,
  });
  const animeId = getOrCreateAnimeRecord(db, {
    parsedTitle: title,
    canonicalTitle: title,
    anilistId: null,
    titleRomaji: null,
    titleEnglish: null,
    titleNative: null,
    metadataJson: null,
  });
  linkVideoToAnimeRecord(db, videoId, {
    animeId,
    parsedBasename: `${key}.mkv`,
    parsedTitle: title,
    parsedSeason: 1,
    parsedEpisode: episode,
    parserSource: 'test',
    parserConfidence: 1,
    parseMetadataJson: null,
  });
  if (watched) db.prepare('UPDATE imm_videos SET watched = 1 WHERE video_id = ?').run(videoId);
  return videoId;
}

function addRollup(
  db: DatabaseSync,
  day: number,
  videoId: number,
  values: { minutes?: number; words?: number; cards?: number } = {},
): void {
  db.prepare(
    `INSERT INTO imm_daily_rollups (
       rollup_day, video_id, total_sessions, total_active_min, total_lines_seen,
       total_tokens_seen, total_cards
     ) VALUES (?, ?, 1, ?, 1, ?, ?)`,
  ).run(day, videoId, values.minutes ?? 30, values.words ?? 100, values.cards ?? 0);
}

function addSession(
  db: DatabaseSync,
  day: number,
  videoId: number,
  values: { minutes?: number; lookups?: number } = {},
): void {
  const startedAtMs = localNoonMs(day).toString();
  const activeMs = (values.minutes ?? 30) * 60_000;
  db.prepare(
    `INSERT INTO imm_sessions (
       session_uuid, video_id, started_at_ms, ended_at_ms, status, active_watched_ms,
       total_watched_ms, yomitan_lookup_count
     ) VALUES (?, ?, ?, ?, 2, ?, ?, ?)`,
  ).run(
    `session-${day}-${videoId}-${Math.random()}`,
    videoId,
    startedAtMs,
    (BigInt(startedAtMs) + BigInt(activeMs)).toString(),
    activeMs,
    activeMs,
    values.lookups ?? 0,
  );
}

function addWord(
  db: DatabaseSync,
  values: {
    headword: string;
    word?: string;
    day: number;
    pos2?: string | null;
    rank?: number | null;
  },
): void {
  const firstSeen = localNoonMs(values.day) / 1000;
  db.prepare(
    `INSERT INTO imm_words (
       headword, word, reading, part_of_speech, pos1, pos2, pos3, first_seen, last_seen,
       frequency, frequency_rank
     ) VALUES (?, ?, ?, 'noun', '名詞', ?, '', ?, ?, 1, ?)`,
  ).run(
    values.headword,
    values.word ?? values.headword,
    values.word ?? values.headword,
    values.pos2 ?? null,
    firstSeen,
    firstSeen,
    values.rank ?? null,
  );
}

test('the title lists stay inside the calendar range when activity is sparse', () => {
  withDashboardDb((db) => {
    const recent = addVideo(db, 'Recent Show', 1);
    const old = addVideo(db, 'Old Show', 1);
    addRollup(db, TODAY, recent, { minutes: 20 });
    // Outside a 7-day calendar window, though well within the last 7 active days.
    addRollup(db, TODAY - 40, old, { minutes: 90 });
    addSession(db, TODAY - 40, old);

    const dashboard = getTrendsDashboard(db, '7d', 'day');

    assert.deepEqual(
      dashboard.librarySummary.map((row) => row.title),
      ['Recent Show'],
    );
    assert.deepEqual(
      dashboard.titleTimeline.map((point) => point.animeTitle),
      ['Recent Show'],
    );
    assert.equal(dashboard.summary.current.watchTimeMin, 20);
    assert.equal(
      dashboard.librarySummary.reduce((sum, row) => sum + row.watchTimeMin, 0),
      dashboard.activity.watchTime.reduce((sum, point) => sum + point.value, 0),
    );
  });
});

test('every grouping reports the same totals for a range', () => {
  withDashboardDb((db) => {
    const videoId = addVideo(db, 'Show', 1);
    for (const day of [TODAY, TODAY - 3, TODAY - 20, TODAY - 29, TODAY - 45]) {
      addRollup(db, day, videoId, { minutes: 10, cards: 1 });
    }
    // January is only partly in range, and its monthly total includes days that are not.
    const insertMonthly = db.prepare(
      `INSERT INTO imm_monthly_rollups (
         rollup_month, video_id, total_sessions, total_active_min, total_lines_seen,
         total_tokens_seen, total_cards
       ) VALUES (?, ?, 1, ?, 1, 100, 1)`,
    );
    insertMonthly.run(202601, videoId, 20);
    insertMonthly.run(202602, videoId, 20);
    insertMonthly.run(202603, videoId, 10);

    const sum = (groupBy: 'day' | 'week' | 'month') =>
      getTrendsDashboard(db, '30d', groupBy).activity.watchTime.reduce(
        (total, point) => total + point.value,
        0,
      );

    assert.equal(sum('day'), 40);
    assert.equal(sum('week'), 40);
    assert.equal(sum('month'), 40);
  });
});

test('all-time month buckets use monthly rollups that outlive daily ones', () => {
  withDashboardDb((db) => {
    const videoId = addVideo(db, 'Show', 1);
    const insertMonthly = db.prepare(
      `INSERT INTO imm_monthly_rollups (
         rollup_month, video_id, total_sessions, total_active_min, total_lines_seen,
         total_tokens_seen, total_cards
       ) VALUES (?, ?, 3, ?, 1, 100, 50)`,
    );
    insertMonthly.run(202501, videoId, 500);
    insertMonthly.run(202503, videoId, 200);

    assert.deepEqual(getTrendsDashboard(db, 'all', 'month').activity.watchTime, [
      { label: 'Jan 2025', value: 500 },
      { label: 'Feb 2025', value: 0 },
      { label: 'Mar 2025', value: 200 },
    ]);
    assert.deepEqual(getTrendsDashboard(db, 'all', 'month', false).activity.watchTime, [
      { label: 'Jan 2025', value: 500 },
      { label: 'Mar 2025', value: 200 },
    ]);
  });
});

test('a finished video counts once however many days it was watched', () => {
  withDashboardDb((db) => {
    const finished = addVideo(db, 'Show', 1, true);
    const unfinished = addVideo(db, 'Show', 2, false);
    addRollup(db, TODAY - 2, finished);
    addRollup(db, TODAY - 1, finished);
    addRollup(db, TODAY, unfinished);

    for (const groupBy of ['day', 'week', 'month'] as const) {
      const dashboard = getTrendsDashboard(db, '30d', groupBy);
      assert.equal(dashboard.progress.episodes.at(-1)?.value, 1, groupBy);
      assert.equal(dashboard.summary.current.episodes, 1, groupBy);
    }

    const daily = getTrendsDashboard(db, '7d', 'day').activity.episodes;
    // Attributed to the last day it was watched.
    assert.equal(daily.at(-2)?.value, 1);
    assert.equal(daily.at(-3)?.value, 0);
    assert.equal(getTrendsDashboard(db, '30d', 'day').librarySummary[0]?.episodesFinished, 1);
    assert.equal(getTrendsDashboard(db, '30d', 'day').librarySummary[0]?.videos, 2);
  });
});

test('rates are gaps without enough data, not zeros', () => {
  withDashboardDb((db) => {
    const videoId = addVideo(db, 'Show', 1);
    // A two-minute day with one card: 30 cards per hour if rated on its own.
    addRollup(db, TODAY - 20, videoId, { minutes: 2, words: 50, cards: 1 });

    const dashboard = getTrendsDashboard(db, '30d', 'day');

    assert.equal(dashboard.ratios.cardsPerHour.at(-1)?.value, null);
    assert.equal(
      dashboard.ratios.cardsPerHour.every((point) => point.value === null),
      true,
    );
    assert.equal(
      dashboard.ratios.readingSpeed.every((point) => point.value === null),
      true,
    );
  });
});

test('daily rates are volume-weighted over a trailing week, reaching before the range', () => {
  withDashboardDb((db) => {
    const videoId = addVideo(db, 'Show', 1);
    // Eight days before today: outside the 7d range, inside the first day's lookback.
    addRollup(db, TODAY - 8, videoId, { minutes: 60, words: 3_000, cards: 6 });
    addRollup(db, TODAY - 6, videoId, { minutes: 60, words: 1_000, cards: 2 });

    const speed = getTrendsDashboard(db, '7d', 'day').ratios.readingSpeed;

    // (3000 + 1000) words / (60 + 60) min, not an average of 50 and 16.7 ...
    assert.equal(speed[0]?.value, +(4_000 / 120).toFixed(1));
    // ... and once the older day leaves the trailing week only the second remains.
    assert.equal(speed[5]?.value, +(1_000 / 60).toFixed(1));
  });
});

test('week grouping starts on Monday and clips the first bucket to the range', () => {
  withDashboardDb((db) => {
    const videoId = addVideo(db, 'Show', 1);
    addRollup(db, TODAY, videoId);

    const { buckets } = getTrendsDashboard(db, '30d', 'week');

    assert.equal(buckets[0]?.startDay, TODAY - 29);
    assert.equal(buckets.at(-1)?.endDay, TODAY);
    // Today is a Sunday, so the final bucket is a full Monday-to-Sunday week.
    assert.equal(buckets.at(-1)?.startDay, TODAY - 6);
    for (let index = 1; index < buckets.length; index += 1) {
      assert.equal(buckets[index]!.startDay, buckets[index - 1]!.endDay + 1);
    }
  });
});

test('showing empty periods on all-time fills the calendar between active days', () => {
  withDashboardDb((db) => {
    const videoId = addVideo(db, 'Show', 1);
    addRollup(db, TODAY - 10, videoId);
    addRollup(db, TODAY, videoId);

    assert.equal(getTrendsDashboard(db, 'all', 'day', true).activity.watchTime.length, 11);
    assert.equal(getTrendsDashboard(db, 'all', 'day', false).activity.watchTime.length, 2);
    assert.deepEqual(getTrendsDashboard(db, 'all', 'day', true).window, {
      startDay: TODAY - 10,
      endDay: TODAY,
    });
  });
});

test('month labels name the year, and day labels do when the range crosses one', () => {
  withDashboardDb((db) => {
    const videoId = addVideo(db, 'Show', 1);
    addRollup(db, TODAY, videoId);

    assert.equal(getTrendsDashboard(db, '30d', 'month').buckets.at(-1)?.label, 'Mar 2026');
    assert.equal(getTrendsDashboard(db, '7d', 'day').buckets.at(-1)?.label, 'Mar 1');
    assert.equal(getTrendsDashboard(db, '90d', 'day').buckets.at(-1)?.label, 'Mar 1, 2026');
  });
});

test('a custom range clips to its days and compares against the same-length span before it', () => {
  withDashboardDb((db) => {
    const videoId = addVideo(db, 'Show', 1);
    addRollup(db, TODAY - 5, videoId, { minutes: 10 });
    addRollup(db, TODAY - 15, videoId, { minutes: 20 });
    addRollup(db, TODAY - 25, videoId, { minutes: 40 });

    const dashboard = getTrendsDashboard(db, { startDay: TODAY - 19, endDay: TODAY - 10 });

    assert.equal(dashboard.activity.watchTime.length, 10);
    assert.equal(dashboard.summary.current.watchTimeMin, 20);
    assert.equal(dashboard.summary.previous?.watchTimeMin, 40);
    assert.deepEqual(dashboard.window, { startDay: TODAY - 19, endDay: TODAY - 10 });
  });
});

test('summary compares with the previous period and all-time has no comparison', () => {
  withDashboardDb((db) => {
    const videoId = addVideo(db, 'Show', 1);
    addRollup(db, TODAY, videoId, { minutes: 10, cards: 1 });
    addRollup(db, TODAY - 10, videoId, { minutes: 25, cards: 4 });

    const { current, previous } = getTrendsDashboard(db, '7d', 'day').summary;

    assert.equal(current.watchTimeMin, 10);
    assert.equal(current.activeDays, 1);
    assert.equal(previous?.watchTimeMin, 25);
    assert.equal(previous?.cards, 4);
    assert.equal(getTrendsDashboard(db, 'all', 'day').summary.previous, null);
  });
});

test('a title filter narrows the page but keeps the title lists complete', () => {
  withDashboardDb((db) => {
    const first = addVideo(db, 'First', 1);
    const second = addVideo(db, 'Second', 1);
    addRollup(db, TODAY, first, { minutes: 10 });
    addRollup(db, TODAY, second, { minutes: 50 });
    addSession(db, TODAY, first, { minutes: 10, lookups: 3 });
    addSession(db, TODAY, second, { minutes: 50, lookups: 9 });
    addWord(db, { headword: '海', day: TODAY });

    const dashboard = getTrendsDashboard(db, '7d', 'day', true, 'First');

    assert.equal(dashboard.summary.current.watchTimeMin, 10);
    assert.equal(dashboard.progress.lookups.at(-1)?.value, 3);
    assert.equal(
      dashboard.watchTimeByDayHour.flat().reduce((sum, value) => sum + value, 0),
      10,
    );
    assert.equal(dashboard.librarySummary.length, 2);
    assert.equal(dashboard.titleTimeline.length, 2);
    // New words are not tracked per title.
    assert.equal(dashboard.summary.current.newWords, null);
    assert.equal(dashboard.learning, null);
    assert.deepEqual(dashboard.progress.newWords, []);
  });
});

test('running totals follow the activity charts when empty days are hidden', () => {
  withDashboardDb((db) => {
    const videoId = addVideo(db, 'Show', 1);
    addRollup(db, TODAY - 3, videoId, { minutes: 30 });
    addRollup(db, TODAY, videoId, { minutes: 10 });

    const dashboard = getTrendsDashboard(db, '7d', 'day', false);

    assert.deepEqual(
      dashboard.progress.watchTime.map((point) => point.value),
      [30, 40],
    );
    assert.deepEqual(
      dashboard.progress.watchTime.map((point) => point.label),
      dashboard.activity.watchTime.map((point) => point.label),
    );
  });
});

test('the watch time running total rounds its sum, so it ends on the summary total', () => {
  withDashboardDb((db) => {
    const videoId = addVideo(db, 'Show', 1);
    addRollup(db, TODAY - 2, videoId, { minutes: 10.4 });
    addRollup(db, TODAY - 1, videoId, { minutes: 10.4 });
    addRollup(db, TODAY, videoId, { minutes: 10.4 });

    const dashboard = getTrendsDashboard(db, '7d', 'day', false);

    assert.deepEqual(
      dashboard.activity.watchTime.map((point) => point.value),
      [10, 10, 10],
    );
    assert.deepEqual(
      dashboard.progress.watchTime.map((point) => point.value),
      [10, 21, 31],
    );
    assert.equal(dashboard.summary.current.watchTimeMin, 31);
  });
});

test('the summary adds watch time in date order, so a total near .5 rounds like the chart', () => {
  withDashboardDb((db) => {
    const videoId = addVideo(db, 'Show', 1);
    // Newest first these add up to 4.499999999999999, and oldest first to 4.5.
    addRollup(db, TODAY - 2, videoId, { minutes: 0.1 });
    addRollup(db, TODAY - 1, videoId, { minutes: 0.1 });
    addRollup(db, TODAY, videoId, { minutes: 4.3 });

    const dashboard = getTrendsDashboard(db, '7d', 'day');

    assert.equal(
      dashboard.summary.current.watchTimeMin,
      dashboard.progress.watchTime.at(-1)?.value,
    );
  });
});

test('new words count a dictionary word once, on its first form, without names', () => {
  withDashboardDb((db) => {
    addWord(db, { headword: '言う', word: '言う', day: TODAY - 2, rank: 50 });
    addWord(db, { headword: '言う', word: '言った', day: TODAY, rank: 50 });
    addWord(db, { headword: '東京', day: TODAY - 1, pos2: '固有名詞', rank: 800 });

    const dashboard = getTrendsDashboard(db, '7d', 'day');

    assert.equal(dashboard.summary.current.newWords, 1);
    assert.equal(dashboard.progress.newWords.at(-1)?.value, 1);
    assert.equal(dashboard.progress.newWords.at(-3)?.value, 1);
    assert.equal(dashboard.progress.newWords.at(-4)?.value, 0);
  });
});

test('new words per thousand and frequency bands read dictionary words', () => {
  withDashboardDb((db) => {
    const videoId = addVideo(db, 'Show', 1);
    addRollup(db, TODAY, videoId, { words: 2_000 });
    addWord(db, { headword: '一', day: TODAY, rank: 500 });
    addWord(db, { headword: '二', day: TODAY, rank: 3_000 });
    addWord(db, { headword: '三', day: TODAY, rank: 30_000 });
    addWord(db, { headword: '四', day: TODAY, rank: null });

    const { learning } = getTrendsDashboard(db, '7d', 'day');

    assert.equal(learning?.newWordsPerThousand.at(-1)?.value, 2);
    assert.deepEqual(learning?.frequencyBands.at(-1)?.counts, {
      top1k: 1,
      top5k: 1,
      top10k: 0,
      top20k: 0,
      beyond20k: 1,
      unranked: 1,
    });
  });
});

test('vocabulary charts and summary count dictionary words, not surface forms', () => {
  withDashboardDb((db) => {
    markLexicalDailyRollupsReady(db);
    addWord(db, { headword: '言う', word: '言う', day: TODAY - 20 });
    addWord(db, { headword: '言う', word: '言った', day: TODAY });
    addWord(db, { headword: '東京', day: TODAY, pos2: '固有名詞' });

    const charts = getVocabularyChartData(db);
    const summary = getVocabularySummary(db, null, Number(NOW_MS));

    assert.deepEqual(charts.newWordsTimelineWithoutNames, [{ epochDay: TODAY - 20, wordCount: 1 }]);
    assert.deepEqual(charts.newWordsTimeline, [
      { epochDay: TODAY - 20, wordCount: 1 },
      { epochDay: TODAY, wordCount: 1 },
    ]);
    assert.equal(summary.uniqueWords, 2);
    assert.equal(summary.uniqueWordsWithoutNames, 1);
    // 言う was first seen long ago, so its newer conjugation is not a new word.
    assert.equal(summary.newThisWeek, 1);
    assert.equal(summary.newThisWeekWithoutNames, 0);
  });
});

test('names tagged with merged parts of speech count as names', () => {
  withDashboardDb((db) => {
    markLexicalDailyRollupsReady(db);
    addWord(db, { headword: '美琴', day: TODAY, pos2: '固有名詞|一般' });
    addWord(db, { headword: '谷君', day: TODAY, pos2: '固有名詞|接尾' });
    addWord(db, { headword: '海', day: TODAY, pos2: '一般' });

    assert.equal(getTrendsDashboard(db, '7d', 'day').summary.current.newWords, 1);
    const summary = getVocabularySummary(db, null, Number(NOW_MS));
    assert.equal(summary.uniqueWords, 3);
    assert.equal(summary.uniqueWordsWithoutNames, 1);
    assert.deepEqual(getVocabularyChartData(db).newWordsTimelineWithoutNames, [
      { epochDay: TODAY, wordCount: 1 },
    ]);
  });
});

test('new word counts pick up writes from this and other connections', () => {
  withDashboardDb((db) => {
    addWord(db, { headword: '海', day: TODAY });
    assert.equal(getTrendsDashboard(db, '7d', 'day').summary.current.newWords, 1);

    addWord(db, { headword: '山', day: TODAY });
    assert.equal(getTrendsDashboard(db, '7d', 'day').summary.current.newWords, 2);

    const dbPath = (db.prepare('PRAGMA database_list').get() as { file: string }).file;
    const other = new Database(dbPath);
    try {
      addWord(other, { headword: '川', day: TODAY });
    } finally {
      other.close();
    }
    assert.equal(getTrendsDashboard(db, '7d', 'day').summary.current.newWords, 3);
  });
});

test('getSessionSummaries can narrow to a span of local days', () => {
  withDashboardDb((db) => {
    const videoId = addVideo(db, 'Show', 1);
    addSession(db, TODAY - 5, videoId);
    addSession(db, TODAY - 4, videoId);
    addSession(db, TODAY, videoId);

    assert.equal(getSessionSummaries(db, 50).length, 3);
    assert.equal(getSessionSummaries(db, 50, { startDay: TODAY - 5, endDay: TODAY - 4 }).length, 2);
  });
});

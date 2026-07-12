import test from 'node:test';
import assert from 'node:assert/strict';
import { formatBytes, formatRelativeTime, summarizeMergeCounts } from './syncui-format';

test('formatBytes renders human readable sizes', () => {
  assert.equal(formatBytes(0), '0 B');
  assert.equal(formatBytes(532), '532 B');
  assert.equal(formatBytes(2048), '2.0 KB');
  assert.equal(formatBytes(5 * 1024 * 1024), '5.0 MB');
  assert.equal(formatBytes(1.5 * 1024 * 1024 * 1024), '1.5 GB');
});

test('formatRelativeTime renders past timestamps', () => {
  const now = 1_700_000_000_000;
  assert.equal(formatRelativeTime(null, now), 'never');
  assert.equal(formatRelativeTime(now - 20_000, now), 'just now');
  assert.equal(formatRelativeTime(now - 5 * 60_000, now), '5 min ago');
  assert.equal(formatRelativeTime(now - 3 * 3_600_000, now), '3 h ago');
  assert.equal(formatRelativeTime(now - 2 * 86_400_000, now), '2 d ago');
});

test('summarizeMergeCounts lists only non-zero counts', () => {
  const summary = {
    sessionsMerged: 3,
    sessionsAlreadyPresent: 12,
    activeSessionsSkipped: 0,
    animeAdded: 1,
    videosAdded: 2,
    wordsAdded: 0,
    kanjiAdded: 0,
    subtitleLinesAdded: 450,
    telemetryRowsAdded: 3,
    eventsAdded: 4,
    excludedWordsAdded: 0,
    dailyRollupsCopied: 0,
    monthlyRollupsCopied: 0,
    rollupGroupsRecomputed: 4,
  };
  const lines = summarizeMergeCounts(summary);
  assert.ok(lines.some((line) => line.label === 'Telemetry rows' && line.value === 3));
  assert.ok(lines.some((line) => line.label === 'Events' && line.value === 4));
  assert.deepEqual(lines, [
    { label: 'Sessions merged', value: 3 },
    { label: 'Already present', value: 12 },
    { label: 'Series added', value: 1 },
    { label: 'Videos added', value: 2 },
    { label: 'Subtitle lines', value: 450 },
    { label: 'Telemetry rows', value: 3 },
    { label: 'Events', value: 4 },
    { label: 'Rollups recomputed', value: 4 },
  ]);

  const empty = summarizeMergeCounts({
    ...summary,
    sessionsMerged: 0,
    sessionsAlreadyPresent: 0,
    animeAdded: 0,
    videosAdded: 0,
    subtitleLinesAdded: 0,
    telemetryRowsAdded: 0,
    eventsAdded: 0,
    rollupGroupsRecomputed: 0,
  });
  assert.deepEqual(empty, [{ label: 'Sessions merged', value: 0 }]);
});

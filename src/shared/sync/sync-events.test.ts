import test from 'node:test';
import assert from 'node:assert/strict';
import { parseSyncProgressLine } from './sync-events';

test('parseSyncProgressLine parses known event types', () => {
  assert.deepEqual(
    parseSyncProgressLine('{"type":"stage","stage":"snapshot-local","message":"Snapshotting"}'),
    { type: 'stage', stage: 'snapshot-local', message: 'Snapshotting' },
  );
  const summaryLine = JSON.stringify({
    type: 'merge-summary',
    target: 'local',
    summary: {
      sessionsMerged: 2,
      sessionsAlreadyPresent: 0,
      activeSessionsSkipped: 0,
      animeAdded: 0,
      videosAdded: 0,
      wordsAdded: 0,
      kanjiAdded: 0,
      subtitleLinesAdded: 0,
      telemetryRowsAdded: 0,
      eventsAdded: 0,
      excludedWordsAdded: 0,
      dailyRollupsCopied: 0,
      monthlyRollupsCopied: 0,
      rollupGroupsRecomputed: 0,
    },
  });
  const parsed = parseSyncProgressLine(summaryLine);
  assert.equal(parsed?.type, 'merge-summary');
  assert.deepEqual(parseSyncProgressLine('{"type":"result","ok":true,"error":null}'), {
    type: 'result',
    ok: true,
    error: null,
  });
});

test('parseSyncProgressLine rejects non-events and garbage', () => {
  assert.equal(parseSyncProgressLine('plain text output'), null);
  assert.equal(parseSyncProgressLine('{"no":"type"}'), null);
  assert.equal(parseSyncProgressLine('{"type":"unknown-event"}'), null);
  assert.equal(parseSyncProgressLine(''), null);
  assert.equal(parseSyncProgressLine('42'), null);
});

test('parseSyncProgressLine rejects events missing required fields', () => {
  // A half-formed event used to be cast straight through; the UI reads
  // `.ok` / `.summary` directly and would misreport a failed sync as done.
  assert.equal(parseSyncProgressLine('{"type":"result"}'), null);
  assert.equal(parseSyncProgressLine('{"type":"result","ok":"yes","error":null}'), null);
  assert.equal(parseSyncProgressLine('{"type":"merge-summary","target":"local"}'), null);
  assert.equal(
    parseSyncProgressLine('{"type":"merge-summary","target":"nowhere","summary":{}}'),
    null,
  );
  assert.equal(parseSyncProgressLine('{"type":"stage","stage":"bogus","message":"x"}'), null);
  assert.equal(parseSyncProgressLine('{"type":"snapshot-created"}'), null);
  assert.equal(parseSyncProgressLine('{"type":"check-result","host":"h","sshOk":true}'), null);

  // Well-formed events still parse.
  assert.deepEqual(parseSyncProgressLine('{"type":"result","ok":true,"error":null}'), {
    type: 'result',
    ok: true,
    error: null,
  });
});

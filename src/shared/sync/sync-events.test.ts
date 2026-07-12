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
    summary: { sessionsMerged: 2 },
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

import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import { toErrorMessage } from './useSessions';

const USE_SESSIONS_PATH = fileURLToPath(new URL('./useSessions.ts', import.meta.url));

test('toErrorMessage normalizes Error and non-Error rejections', () => {
  assert.equal(toErrorMessage(new Error('network down')), 'network down');
  assert.equal(toErrorMessage('bad gateway'), 'bad gateway');
  assert.equal(toErrorMessage(503), '503');
});

test('useSessions and useSessionDetail route catch handlers through toErrorMessage', () => {
  const source = fs.readFileSync(USE_SESSIONS_PATH, 'utf8');
  const matches = source.match(/setError\(toErrorMessage\(err\)\)/g);

  assert.equal(matches?.length, 2);
});

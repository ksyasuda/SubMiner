import assert from 'node:assert/strict';
import test from 'node:test';
import { formatSkippedYomitanWriteAction } from './yomitan-read-only-log';

test('formatSkippedYomitanWriteAction redacts full filesystem paths to basenames', () => {
  assert.equal(
    formatSkippedYomitanWriteAction('importYomitanDictionary', '/tmp/private/merged.zip'),
    'importYomitanDictionary(merged.zip)',
  );
});

test('formatSkippedYomitanWriteAction redacts dictionary titles', () => {
  assert.equal(
    formatSkippedYomitanWriteAction('deleteYomitanDictionary', 'SubMiner Character Dictionary'),
    'deleteYomitanDictionary(<redacted>)',
  );
});

test('formatSkippedYomitanWriteAction falls back when value is blank', () => {
  assert.equal(
    formatSkippedYomitanWriteAction('upsertYomitanDictionarySettings', '   '),
    'upsertYomitanDictionarySettings(<redacted>)',
  );
});

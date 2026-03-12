import assert from 'node:assert/strict';
import test from 'node:test';
import {
  getCharacterDictionaryDisabledReason,
  isCharacterDictionaryRuntimeEnabled,
} from './character-dictionary-availability';

test('character dictionary runtime is enabled when external Yomitan profile is not configured', () => {
  assert.equal(isCharacterDictionaryRuntimeEnabled(''), true);
  assert.equal(isCharacterDictionaryRuntimeEnabled('   '), true);
  assert.equal(getCharacterDictionaryDisabledReason(''), null);
});

test('character dictionary runtime is disabled when external Yomitan profile is configured', () => {
  assert.equal(isCharacterDictionaryRuntimeEnabled('/tmp/gsm-profile'), false);
  assert.equal(
    getCharacterDictionaryDisabledReason('/tmp/gsm-profile'),
    'Character dictionary is disabled while yomitan.externalProfilePath is configured.',
  );
});

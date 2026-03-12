import assert from 'node:assert/strict';
import test from 'node:test';
import { createYomitanProfilePolicy } from './yomitan-profile-policy';

test('yomitan profile policy trims external profile path and marks read-only mode', () => {
  const calls: string[] = [];
  const policy = createYomitanProfilePolicy({
    externalProfilePath: '  /tmp/gsm-profile  ',
    logInfo: (message) => calls.push(message),
  });

  assert.equal(policy.externalProfilePath, '/tmp/gsm-profile');
  assert.equal(policy.isExternalReadOnlyMode(), true);
  assert.equal(policy.isCharacterDictionaryEnabled(), false);
  assert.equal(
    policy.getCharacterDictionaryDisabledReason(),
    'Character dictionary is disabled while yomitan.externalProfilePath is configured.',
  );

  policy.logSkippedWrite('importYomitanDictionary(sample.zip)');
  assert.deepEqual(calls, [
    '[yomitan] skipping importYomitanDictionary(sample.zip): yomitan.externalProfilePath is configured; external profile mode is read-only',
  ]);
});

test('yomitan profile policy keeps character dictionary enabled without external profile path', () => {
  const policy = createYomitanProfilePolicy({
    externalProfilePath: '   ',
    logInfo: () => undefined,
  });

  assert.equal(policy.externalProfilePath, '');
  assert.equal(policy.isExternalReadOnlyMode(), false);
  assert.equal(policy.isCharacterDictionaryEnabled(), true);
  assert.equal(policy.getCharacterDictionaryDisabledReason(), null);
});

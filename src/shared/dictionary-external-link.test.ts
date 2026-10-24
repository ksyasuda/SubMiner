import test from 'node:test';
import assert from 'node:assert/strict';
import { parseDictionaryExternalUrl } from './dictionary-external-link';

test('dictionary links accept web URLs and reject privileged schemes and credentials', () => {
  assert.equal(
    parseDictionaryExternalUrl('https://example.com/word?q=猫'),
    'https://example.com/word?q=%E7%8C%AB',
  );
  for (const value of [
    'file:///etc/passwd',
    'javascript:alert(1)',
    'https://user:pass@example.com',
    null,
    {},
  ]) {
    assert.throws(() => parseDictionaryExternalUrl(value));
  }
});

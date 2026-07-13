import assert from 'node:assert/strict';
import test from 'node:test';
import { isKanaChar, isKanaOnlyText } from './token-classification';

test('kana classification excludes the katakana-hiragana double hyphen', () => {
  assert.equal(isKanaChar('゠'), false);
  assert.equal(isKanaOnlyText('゠'), false);
});

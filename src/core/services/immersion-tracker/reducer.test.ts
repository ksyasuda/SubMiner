import test from 'node:test';
import assert from 'node:assert/strict';
import { extractLineVocabulary, isKanji } from './reducer';

test('isKanji follows canonical CJK ranges', () => {
  assert.ok(isKanji('日'));
  assert.ok(isKanji('𠀀'));
  assert.ok(!isKanji('あ'));
  assert.ok(!isKanji('a'));
});

test('extractLineVocabulary returns words and unique kanji', () => {
  const result = extractLineVocabulary('hello 你好 猫');

  assert.equal(result.words.length, 3);
  assert.deepEqual(
    new Set(result.words.map((entry) => `${entry.headword}/${entry.word}`)),
    new Set(['hello/hello', '你好/你好', '猫/猫']),
  );
  assert.equal(result.words.every((entry) => entry.reading === ''), true);
  assert.deepEqual(new Set(result.kanji), new Set(['你', '好', '猫']));
});

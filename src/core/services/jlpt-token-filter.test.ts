import assert from 'node:assert/strict';
import test from 'node:test';
import {
  getIgnoredPos1Entries,
  JLPT_EXCLUDED_TERMS,
  JLPT_IGNORED_MECAB_POS1,
  JLPT_IGNORED_MECAB_POS1_ENTRIES,
  JLPT_IGNORED_MECAB_POS1_LIST,
  shouldIgnoreJlptByTerm,
  shouldIgnoreJlptForMecabPos1,
} from './jlpt-token-filter';

test('shouldIgnoreJlptByTerm matches the excluded JLPT lexical terms', () => {
  assert.equal(shouldIgnoreJlptByTerm('この'), true);
  assert.equal(shouldIgnoreJlptByTerm('そこ'), true);
  assert.equal(shouldIgnoreJlptByTerm('猫'), false);
  assert.deepEqual(Array.from(JLPT_EXCLUDED_TERMS), [
    'この',
    'その',
    'あの',
    'どの',
    'これ',
    'それ',
    'あれ',
    'どれ',
    'ここ',
    'そこ',
    'あそこ',
    'どこ',
    'こと',
    'ああ',
    'ええ',
    'うう',
    'おお',
    'はは',
    'へえ',
    'ふう',
    'ほう',
  ]);
});

test('shouldIgnoreJlptForMecabPos1 matches the exported ignored POS1 list', () => {
  assert.equal(shouldIgnoreJlptForMecabPos1('助詞'), true);
  assert.equal(shouldIgnoreJlptForMecabPos1('名詞'), false);
  assert.deepEqual(JLPT_IGNORED_MECAB_POS1, JLPT_IGNORED_MECAB_POS1_LIST);
  assert.deepEqual(
    JLPT_IGNORED_MECAB_POS1_ENTRIES.map((entry) => entry.pos1),
    JLPT_IGNORED_MECAB_POS1_LIST,
  );
  assert.deepEqual(getIgnoredPos1Entries(), JLPT_IGNORED_MECAB_POS1_ENTRIES);
});

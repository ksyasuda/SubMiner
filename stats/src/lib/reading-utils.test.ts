import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { fullReading } from './reading-utils';

describe('fullReading', () => {
  it('prefixes leading hiragana from headword', () => {
    // お前 with reading まえ → おまえ
    assert.equal(fullReading('お前', 'まえ'), 'おまえ');
  });

  it('handles katakana stored readings', () => {
    // お前 with katakana reading マエ → おまえ
    assert.equal(fullReading('お前', 'マエ'), 'おまえ');
  });

  it('returns stored reading when it already includes leading kana', () => {
    // Reading already correct
    assert.equal(fullReading('お前', 'おまえ'), 'おまえ');
  });

  it('handles trailing hiragana', () => {
    // 隠す with reading かくす — す is trailing hiragana
    assert.equal(fullReading('隠す', 'かくす'), 'かくす');
  });

  it('handles pure kanji headwords', () => {
    assert.equal(fullReading('様', 'さま'), 'さま');
  });

  it('returns empty for empty reading', () => {
    assert.equal(fullReading('前', ''), '');
  });

  it('returns empty for empty headword', () => {
    assert.equal(fullReading('', 'まえ'), 'まえ');
  });

  it('handles all-kana headword', () => {
    // Headword is already all hiragana
    assert.equal(fullReading('いますぐ', 'いますぐ'), 'いますぐ');
  });

  it('handles mixed leading and trailing kana', () => {
    // お気に入り: お=leading, に入り=trailing around 気
    assert.equal(fullReading('お気に入り', 'きにいり'), 'おきにいり');
  });

  it('handles katakana in headword', () => {
    // カズマ様 — leading katakana + kanji
    assert.equal(fullReading('カズマ様', 'さま'), 'かずまさま');
  });
});

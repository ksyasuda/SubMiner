import { describe, it, expect } from 'vitest';
import { fullReading } from './reading-utils';

describe('fullReading', () => {
  it('prefixes leading hiragana from headword', () => {
    // お前 with reading まえ → おまえ
    expect(fullReading('お前', 'まえ')).toBe('おまえ');
  });

  it('handles katakana stored readings', () => {
    // お前 with katakana reading マエ → おまえ
    expect(fullReading('お前', 'マエ')).toBe('おまえ');
  });

  it('returns stored reading when it already includes leading kana', () => {
    // Reading already correct
    expect(fullReading('お前', 'おまえ')).toBe('おまえ');
  });

  it('handles trailing hiragana', () => {
    // 隠す with reading かくす — す is trailing hiragana
    expect(fullReading('隠す', 'かくす')).toBe('かくす');
  });

  it('handles pure kanji headwords', () => {
    expect(fullReading('様', 'さま')).toBe('さま');
  });

  it('returns empty for empty reading', () => {
    expect(fullReading('前', '')).toBe('');
  });

  it('returns empty for empty headword', () => {
    expect(fullReading('', 'まえ')).toBe('まえ');
  });

  it('handles all-kana headword', () => {
    // Headword is already all hiragana
    expect(fullReading('いますぐ', 'いますぐ')).toBe('いますぐ');
  });

  it('handles mixed leading and trailing kana', () => {
    // お気に入り: お=leading, に入り=trailing around 気
    expect(fullReading('お気に入り', 'きにいり')).toBe('おきにいり');
  });

  it('handles katakana in headword', () => {
    // カズマ様 — leading katakana + kanji
    expect(fullReading('カズマ様', 'さま')).toBe('かずまさま');
  });
});

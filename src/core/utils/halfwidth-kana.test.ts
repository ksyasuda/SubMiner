import assert from 'node:assert/strict';
import test from 'node:test';
import { narrowFullwidthKana, widenHalfwidthKana } from './halfwidth-kana';

test('widenHalfwidthKana widens halfwidth katakana and composes voiced marks', () => {
  assert.equal(
    widenHalfwidthKana('(ﾄｰｶ)ﾀﾞｷﾆお姉ちゃん ｱﾝﾀがｱｰﾀﾙたん｡'),
    '(トーカ)ダキニお姉ちゃん アンタがアータルたん。',
  );
});

test('widenHalfwidthKana leaves other compatibility characters alone', () => {
  assert.equal(widenHalfwidthKana('見て見てっ！①㍑ＡＢ'), '見て見てっ！①㍑ＡＢ');
});

test('widenHalfwidthKana folds a lone halfwidth voiced mark into the kana before it', () => {
  assert.equal(widenHalfwidthKana('山カﾞク ハﾟン'), '山ガク パン');
  assert.equal(widenHalfwidthKana('アﾞ'), 'ア\u3099');
});

test('narrowFullwidthKana inverts widening and leaves other text alone', () => {
  assert.equal(narrowFullwidthKana('ダキニお姉ちゃん アンタ。ヴ'), 'ﾀﾞｷﾆお姉ちゃん ｱﾝﾀ｡ｳﾞ');
  assert.equal(narrowFullwidthKana('漢字ひらがな！'), '漢字ひらがな！');
});

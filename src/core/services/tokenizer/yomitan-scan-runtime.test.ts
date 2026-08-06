import assert from 'node:assert/strict';
import test from 'node:test';
import { requestYomitanScanTokens } from './yomitan-parser-runtime';
import {
  countTermsFindLookups,
  createNameScanDeps,
  NAME_SCAN_WORDS,
} from './yomitan-scan-test-harness';

// Behaviour of the in-page scan runtime around character names and kana:
// which positions the greedy pre-pass probes, and what the walk makes of
// halfwidth spellings. Driven end to end through requestYomitanScanTokens
// because the runtime only exists inside the parser window.

const NAME_SCAN_LINE = 'ミナトはまだ学校にいない';

test('requestYomitanScanTokens skips name pre-pass lookups where no candidate name can start', async () => {
  const exhaustiveLookups: string[] = [];
  const exhaustive = await requestYomitanScanTokens(
    NAME_SCAN_LINE,
    createNameScanDeps(exhaustiveLookups),
    { error: () => undefined },
    { includeNameMatchMetadata: true },
  );

  const prefilteredLookups: string[] = [];
  const prefiltered = await requestYomitanScanTokens(
    NAME_SCAN_LINE,
    createNameScanDeps(prefilteredLookups),
    { error: () => undefined },
    {
      includeNameMatchMetadata: true,
      currentCharacterDictionaryMediaId: 1,
      // Terms and readings the generated dictionary exposes for this media.
      nameCandidates: { key: 'media-1', forms: ['ミナト', 'みなと'] },
    },
  );

  // Same tokenization, including the name match, with fewer round trips.
  assert.deepEqual(prefiltered, exhaustive);
  assert.equal(prefiltered?.[0]?.surface, 'ミナト');
  assert.equal(prefiltered?.[0]?.isNameMatch, true);
  assert.ok(
    prefilteredLookups.length < exhaustiveLookups.length,
    `expected fewer lookups with candidates (${prefilteredLookups.length} vs ${exhaustiveLookups.length})`,
  );
  // Mid-token positions are exactly what the pre-pass used to probe (a name can
  // start mid-token); with candidates they cost nothing, while the main walk's
  // own token-start lookups are unaffected.
  assert.ok(countTermsFindLookups(exhaustiveLookups, '校に') > 0);
  assert.equal(countTermsFindLookups(prefilteredLookups, '校に'), 0);
});

test('requestYomitanScanTokens matches a katakana name from its kana-normalized candidate form', async () => {
  const lookups: string[] = [];
  const result = await requestYomitanScanTokens(
    NAME_SCAN_LINE,
    createNameScanDeps(lookups),
    { error: () => undefined },
    {
      includeNameMatchMetadata: true,
      currentCharacterDictionaryMediaId: 1,
      // Only the hiragana reading is listed; the katakana surface in the line
      // must still be found through kana normalization.
      nameCandidates: { key: 'media-1', forms: ['みなと'] },
    },
  );

  assert.equal(result?.[0]?.surface, 'ミナト');
  assert.equal(result?.[0]?.isNameMatch, true);
});

// Kana normalization folds halfwidth katakana, so a name written that way does
// prefix-match a candidate form — but only if the position counts as Japanese
// in the first place. The generic word here reaches into the name, so only a
// pre-pass reservation can keep the name whole.
const HALFWIDTH_NAME_SCAN_WORDS: Array<[string, string, string, boolean]> = [
  ['ﾈｺ', 'ネコ', 'ねこ', false],
  ['まだﾐ', 'まだミ', 'まだみ', false],
  ['まだ', 'まだ', 'まだ', false],
  ['ﾐﾅﾄ', 'ミナト', 'みなと', true],
];

test('requestYomitanScanTokens probes halfwidth katakana positions during the name pre-pass', async () => {
  const lookups: string[] = [];
  const result = await requestYomitanScanTokens(
    'ﾈｺまだﾐﾅﾄ',
    createNameScanDeps(lookups, HALFWIDTH_NAME_SCAN_WORDS),
    { error: () => undefined },
    {
      includeNameMatchMetadata: true,
      currentCharacterDictionaryMediaId: 1,
      // Fullwidth forms only, as the generated dictionary stores them.
      nameCandidates: { key: 'media-1', forms: ['ミナト', 'みなと'] },
    },
  );

  assert.equal(countTermsFindLookups(lookups, 'ﾐﾅﾄ'), 1);
  // ｺ is mid-token, so only the pre-pass would ever look it up, and it matches
  // no candidate: folding halfwidth made those positions indexable, so they no
  // longer cost a round trip apiece.
  assert.equal(countTermsFindLookups(lookups, 'ｺ'), 0);
  assert.deepEqual(
    result?.map((token) => token.surface),
    ['ﾈｺ', 'まだ', 'ﾐﾅﾄ'],
  );
  assert.equal(result?.[2]?.isNameMatch, true);
  // The reading is written the way the fullwidth katakana path writes it
  // (surface spelling, fullwidth): halfwidth kana is not kana to the known-word
  // and frequency code downstream, and an empty reading there disables the
  // reading fallback entirely.
  assert.equal(result?.[2]?.reading, 'ミナト');
  assert.equal(result?.[2]?.headwordReading, 'みなと');
});

test('a voiced halfwidth name still bypasses the candidate prefilter', async () => {
  const lookups: string[] = [];
  const result = await requestYomitanScanTokens(
    'まだｶﾞｸ',
    createNameScanDeps(lookups, [
      ['まだｶ', 'まだカ', 'まだか', false],
      ['まだ', 'まだ', 'まだ', false],
      ['ｶﾞｸ', 'ガク', 'がく', true],
    ]),
    { error: () => undefined },
    {
      includeNameMatchMetadata: true,
      currentCharacterDictionaryMediaId: 1,
      nameCandidates: { key: 'media-1', forms: ['ガク', 'がく'] },
    },
  );

  // ｶ + ﾞ folds to か + ﾞ, which cannot prefix-match が, so the prefilter would
  // drop this position; the voiced-mark bypass is what keeps the name.
  assert.deepEqual(
    result?.map((token) => token.surface),
    ['まだ', 'ｶﾞｸ'],
  );
  assert.equal(result?.[1]?.isNameMatch, true);
});

test('a mixed-width voiced name survives the candidate prefilter', async () => {
  const lookups: string[] = [];
  const result = await requestYomitanScanTokens(
    'まだ山ｶﾞｸ',
    createNameScanDeps(lookups, [
      ['まだ山', 'まだ山', 'まだやま', false],
      ['まだ', 'まだ', 'まだ', false],
      ['山ｶﾞｸ', '山ガク', 'やまがく', true],
    ]),
    { error: () => undefined },
    {
      includeNameMatchMetadata: true,
      currentCharacterDictionaryMediaId: 1,
      nameCandidates: { key: 'media-1', forms: ['山ガク', 'やまがく'] },
    },
  );

  // The name starts on a kanji, so a bypass keyed on the first character misses
  // it: 山ｶﾞｸ normalizes to 山かﾞく, which cannot match the candidate 山がく, and
  // the generic word starting earlier then swallows the 山.
  assert.deepEqual(
    result?.map((token) => token.surface),
    ['まだ', '山ｶﾞｸ'],
  );
  assert.equal(result?.[1]?.isNameMatch, true);
});

test('halfwidth voiced kana compose into the reading instead of leaving a stray mark', async () => {
  const lookups: string[] = [];
  const result = await requestYomitanScanTokens(
    'ｶﾞｸ ﾊﾟﾝ',
    createNameScanDeps(lookups, [
      ['ｶﾞｸ', 'ガク', 'がく', false],
      ['ﾊﾟﾝ', 'パン', 'ぱん', false],
    ]),
    { error: () => undefined },
    { includeNameMatchMetadata: true },
  );

  // The name pre-pass runs over every position here (no candidate list), but a
  // standalone voiced mark can never start a name, so it costs no lookup.
  assert.equal(countTermsFindLookups(lookups, 'ﾞ'), 0);
  assert.equal(countTermsFindLookups(lookups, 'ﾟ'), 0);
  const readings = (result ?? [])
    .filter((token) => token.isUnparsedRun !== true)
    .map((token) => [token.surface, token.reading]);
  assert.deepEqual(readings, [
    ['ｶﾞｸ', 'ガク'],
    ['ﾊﾟﾝ', 'パン'],
  ]);
});

test('requestYomitanScanTokens falls back to the exhaustive name scan without candidates', async () => {
  const withoutLookups: string[] = [];
  const withoutCandidates = await requestYomitanScanTokens(
    NAME_SCAN_LINE,
    createNameScanDeps(withoutLookups),
    { error: () => undefined },
    { includeNameMatchMetadata: true, currentCharacterDictionaryMediaId: 1, nameCandidates: null },
  );

  assert.equal(withoutCandidates?.[0]?.isNameMatch, true);
  // No candidate list means every Japanese position is probed, as before.
  assert.ok(countTermsFindLookups(withoutLookups, '校に') > 0);
});

test('requestYomitanScanTokens reinstalls name candidates when the media changes', async () => {
  const lookups: string[] = [];
  const deps = createNameScanDeps(lookups);

  // First media's candidates cannot match this line's name.
  const otherMedia = await requestYomitanScanTokens(
    NAME_SCAN_LINE,
    deps,
    { error: () => undefined },
    {
      includeNameMatchMetadata: true,
      currentCharacterDictionaryMediaId: 2,
      nameCandidates: { key: 'media-2', forms: ['カズマ'] },
    },
  );
  assert.equal(otherMedia?.[0]?.isNameMatch, undefined);

  const correctMedia = await requestYomitanScanTokens(
    NAME_SCAN_LINE,
    deps,
    { error: () => undefined },
    {
      includeNameMatchMetadata: true,
      currentCharacterDictionaryMediaId: 1,
      nameCandidates: { key: 'media-1', forms: ['ミナト'] },
    },
  );
  assert.equal(correctMedia?.[0]?.surface, 'ミナト');
  assert.equal(correctMedia?.[0]?.isNameMatch, true);
});

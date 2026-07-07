import assert from 'node:assert/strict';
import test from 'node:test';
import { MergedToken, PartOfSpeech } from '../../../types';
import {
  annotateTokens,
  AnnotationStageDeps,
  shouldExcludeTokenFromSubtitleAnnotations,
  shouldExcludeTokenFromVocabularyPersistence,
  stripSubtitleAnnotationMetadata,
} from './annotation-stage';

function makeToken(overrides: Partial<MergedToken> = {}): MergedToken {
  return {
    surface: '猫',
    reading: 'ネコ',
    headword: '猫',
    startPos: 0,
    endPos: 1,
    partOfSpeech: PartOfSpeech.noun,
    isMerged: false,
    isKnown: false,
    isNPlusOneTarget: false,
    ...overrides,
  };
}

function makeDeps(overrides: Partial<AnnotationStageDeps> = {}): AnnotationStageDeps {
  return {
    isKnownWord: () => false,
    knownWordMatchMode: 'headword',
    getJlptLevel: () => null,
    ...overrides,
  };
}

test('annotateTokens keeps name matches on tokens the POS noise filter would strip', () => {
  // MeCab tags 平 as 接頭詞 in contexts like あっ 平 これ…, which is in the
  // POS1 exclusion list; a confirmed character-name match must survive it.
  const tokens = [
    makeToken({
      surface: '平',
      headword: '平',
      reading: 'たいら',
      pos1: '接頭詞',
      isNameMatch: true,
    }),
    makeToken({
      surface: '平',
      headword: '平',
      reading: 'ひら',
      pos1: '接頭詞',
      isNameMatch: false,
      jlptLevel: 'N1',
    }),
  ];

  const result = annotateTokens(tokens, makeDeps(), { nameMatchEnabled: true });

  assert.equal(result[0]?.isNameMatch, true);
  assert.equal(result[1]?.isNameMatch, false);
  assert.equal(result[1]?.jlptLevel, undefined);
});

test('annotateTokens strips name matches from POS-excluded tokens when name matching is disabled', () => {
  const tokens = [
    makeToken({
      surface: '平',
      headword: '平',
      reading: 'たいら',
      pos1: '接頭詞',
      isNameMatch: true,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps(), { nameMatchEnabled: false });

  assert.equal(result[0]?.isNameMatch, false);
});

test('annotateTokens known-word match mode uses headword vs surface', () => {
  const tokens = [makeToken({ surface: '食べた', headword: '食べる', reading: 'タベタ' })];
  const isKnownWord = (text: string): boolean => text === '食べる';

  const headwordResult = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord,
      knownWordMatchMode: 'headword',
    }),
  );
  const surfaceResult = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord,
      knownWordMatchMode: 'surface',
    }),
  );

  assert.equal(headwordResult[0]?.isKnown, true);
  assert.equal(surfaceResult[0]?.isKnown, false);
});

test('annotateTokens passes dictionary-form reading so spelling collisions stay unknown', () => {
  // とこ (colloquial ところ) resolves to headword 床/とこ; a known 床/ゆか card
  // must not mark it known (#138 regression).
  const cache = new Map([['床', 'ゆか']]);
  const isKnownWord = (text: string, reading?: string): boolean => {
    if (!cache.has(text)) {
      return false;
    }
    return reading === undefined || cache.get(text) === reading;
  };
  const tokens = [
    makeToken({
      surface: 'とこ',
      headword: '床',
      reading: 'とこ',
      headwordReading: 'とこ',
      endPos: 2,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps({ isKnownWord }));

  assert.equal(result[0]?.isKnown, false);
});

test('annotateTokens keeps inflected known words matched via headword reading', () => {
  const isKnownWord = (text: string, reading?: string): boolean =>
    text === '行く' && (reading === undefined || reading === 'いく');
  const tokens = [
    makeToken({
      surface: '行きたい',
      headword: '行く',
      reading: 'いきたい',
      headwordReading: 'いく',
      partOfSpeech: PartOfSpeech.verb,
      endPos: 4,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps({ isKnownWord }));

  assert.equal(result[0]?.isKnown, true);
});

test('annotateTokens omits reading for headword match when token lacks headword reading and is inflected', () => {
  // MeCab tokens have no dictionary-form reading; the surface reading of an
  // inflected form must not be compared against the note's dictionary reading.
  const isKnownWord = (text: string, reading?: string): boolean =>
    text === '食べる' && reading === undefined;
  const tokens = [
    makeToken({
      surface: '食べた',
      headword: '食べる',
      reading: 'タベタ',
      partOfSpeech: PartOfSpeech.verb,
      endPos: 3,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps({ isKnownWord }));

  assert.equal(result[0]?.isKnown, true);
});

test('annotateTokens marks known words when N+1 is disabled', () => {
  const tokens = [
    makeToken({ surface: '私', headword: '私', startPos: 0, endPos: 1 }),
    makeToken({ surface: '猫', headword: '猫', startPos: 1, endPos: 2 }),
    makeToken({ surface: '犬', headword: '犬', startPos: 2, endPos: 3 }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === '私' || text === '猫',
    }),
    { nPlusOneEnabled: false, knownWordsEnabled: true },
  );

  assert.equal(result[0]?.isKnown, true);
  assert.equal(result[0]?.isNPlusOneTarget, false);
  assert.equal(result[1]?.isKnown, true);
  assert.equal(result[1]?.isNPlusOneTarget, false);
  assert.equal(result[2]?.isKnown, false);
  assert.equal(result[2]?.isNPlusOneTarget, false);
});

test('annotateTokens hides known-word marks while still using known words for N+1', () => {
  const tokens = [
    makeToken({ surface: '私', headword: '私', startPos: 0, endPos: 1 }),
    makeToken({ surface: '猫', headword: '猫', startPos: 1, endPos: 2 }),
    makeToken({ surface: '犬', headword: '犬', startPos: 2, endPos: 3 }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === '私' || text === '猫',
    }),
    { nPlusOneEnabled: true, knownWordsEnabled: false, minSentenceWordsForNPlusOne: 3 },
  );

  assert.equal(result[0]?.isKnown, false);
  assert.equal(result[1]?.isKnown, false);
  assert.equal(result[2]?.isKnown, false);
  assert.equal(result[2]?.isNPlusOneTarget, true);
});

test('annotateTokens falls back to reading for known-word matches when headword lookup misses', () => {
  const tokens = [
    makeToken({
      surface: '大体',
      headword: '大体',
      reading: 'だいたい',
      frequencyRank: 1895,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === 'だいたい',
      getJlptLevel: (text) => (text === '大体' ? 'N4' : null),
    }),
  );

  assert.equal(result[0]?.isKnown, true);
  assert.equal(result[0]?.jlptLevel, 'N4');
  assert.equal(result[0]?.frequencyRank, 1895);
});

test('annotateTokens ignores partial furigana readings for known-word fallback', () => {
  const tokens = [
    makeToken({
      surface: '待ち合わせてる',
      headword: '待ち合わせる',
      reading: 'まあ',
      partOfSpeech: PartOfSpeech.verb,
      endPos: 7,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === 'まあ',
    }),
  );

  assert.equal(result[0]?.isKnown, false);
});

test('annotateTokens reading fallback still matches kana surfaces with complete readings', () => {
  const tokens = [
    makeToken({
      surface: 'ください',
      headword: '下さい',
      reading: 'ください',
      partOfSpeech: PartOfSpeech.verb,
      endPos: 4,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === 'ください',
    }),
  );

  assert.equal(result[0]?.isKnown, true);
});

test('annotateTokens excludes frequency for particle/bound_auxiliary and pos1 exclusions', () => {
  const tokens = [
    makeToken({
      surface: 'は',
      headword: 'は',
      partOfSpeech: PartOfSpeech.particle,
      frequencyRank: 3,
    }),
    makeToken({
      surface: 'です',
      headword: 'です',
      partOfSpeech: PartOfSpeech.bound_auxiliary,
      startPos: 1,
      endPos: 3,
      frequencyRank: 4,
    }),
    makeToken({
      surface: 'の',
      headword: 'の',
      partOfSpeech: PartOfSpeech.other,
      pos1: '助詞',
      startPos: 3,
      endPos: 4,
      frequencyRank: 5,
    }),
    makeToken({
      surface: '猫',
      headword: '猫',
      partOfSpeech: PartOfSpeech.noun,
      startPos: 4,
      endPos: 5,
      frequencyRank: 11,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps());

  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[1]?.frequencyRank, undefined);
  assert.equal(result[2]?.frequencyRank, undefined);
  assert.equal(result[3]?.frequencyRank, 11);
});

test('annotateTokens keeps frequency for determiner-led content noun compounds', () => {
  const tokens = [
    makeToken({
      surface: 'その場',
      headword: 'その場',
      reading: 'そのば',
      partOfSpeech: PartOfSpeech.noun,
      pos1: '連体詞|名詞',
      pos2: '*|一般',
      startPos: 0,
      endPos: 3,
      frequencyRank: 879,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === 'その場',
      getJlptLevel: (text) => (text === 'その場' ? 'N4' : null),
    }),
    { minSentenceWordsForNPlusOne: 1 },
  );

  assert.equal(result[0]?.isKnown, true);
  assert.equal(result[0]?.frequencyRank, 879);
  assert.equal(result[0]?.jlptLevel, 'N4');
});

test('annotateTokens preserves existing frequency rank when frequency is enabled', () => {
  const tokens = [makeToken({ surface: '猫', headword: '猫', frequencyRank: 42 })];

  const result = annotateTokens(tokens, makeDeps());

  assert.equal(result[0]?.frequencyRank, 42);
});

test('annotateTokens drops invalid frequency rank values', () => {
  const tokens = [makeToken({ surface: '猫', headword: '猫', frequencyRank: Number.NaN })];
  const result = annotateTokens(tokens, makeDeps());
  assert.equal(result[0]?.frequencyRank, undefined);
});

test('annotateTokens clears frequency rank when frequency is disabled', () => {
  const tokens = [makeToken({ surface: '猫', headword: '猫', frequencyRank: 42 })];
  const result = annotateTokens(tokens, makeDeps(), { frequencyEnabled: false });
  assert.equal(result[0]?.frequencyRank, undefined);
});

test('annotateTokens handles JLPT disabled and eligibility exclusion paths', () => {
  let disabledLookupCalls = 0;
  const disabledResult = annotateTokens(
    [makeToken({ surface: '猫', headword: '猫' })],
    makeDeps({
      getJlptLevel: () => {
        disabledLookupCalls += 1;
        return 'N5';
      },
    }),
    { jlptEnabled: false },
  );
  assert.equal(disabledResult[0]?.jlptLevel, undefined);
  assert.equal(disabledLookupCalls, 0);

  let excludedLookupCalls = 0;
  const excludedResult = annotateTokens(
    [
      makeToken({
        surface: '！',
        headword: '！',
        reading: '',
        pos1: '記号',
        partOfSpeech: PartOfSpeech.symbol,
      }),
    ],
    makeDeps({
      getJlptLevel: () => {
        excludedLookupCalls += 1;
        return 'N5';
      },
    }),
  );
  assert.equal(excludedResult[0]?.jlptLevel, undefined);
  assert.equal(excludedLookupCalls, 0);
});

test('shouldExcludeTokenFromSubtitleAnnotations excludes explanatory ending variants', () => {
  const tokens = [
    makeToken({
      surface: 'んです',
      headword: 'ん',
      reading: 'ンデス',
      pos1: '名詞|助動詞',
      pos2: '非自立',
    }),
    makeToken({
      surface: 'のだ',
      headword: 'の',
      reading: 'ノダ',
      pos1: '名詞|助動詞',
      pos2: '非自立',
    }),
    makeToken({
      surface: 'んだ',
      headword: 'ん',
      reading: 'ンダ',
      pos1: '名詞|助動詞',
      pos2: '非自立',
    }),
    makeToken({
      surface: 'のです',
      headword: 'の',
      reading: 'ノデス',
      pos1: '名詞|助動詞',
      pos2: '非自立',
    }),
    makeToken({
      surface: 'なんです',
      headword: 'だ',
      reading: 'ナンデス',
      pos1: '助動詞|名詞|助動詞',
      pos2: '|非自立',
    }),
    makeToken({
      surface: 'んでした',
      headword: 'ん',
      reading: 'ンデシタ',
      pos1: '助動詞|助動詞|助動詞',
    }),
    makeToken({
      surface: 'のでは',
      headword: 'の',
      reading: 'ノデハ',
      pos1: '助詞|接続詞',
    }),
  ];

  for (const token of tokens) {
    assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true, token.surface);
  }
});

test('shouldExcludeTokenFromSubtitleAnnotations excludes explanatory pondering endings', () => {
  const token = makeToken({
    surface: 'のかな',
    headword: 'の',
    reading: 'ノカナ',
    pos1: '名詞|助動詞',
    pos2: '非自立',
  });

  assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true);
});

test('shouldExcludeTokenFromSubtitleAnnotations excludes explanatory contrast endings', () => {
  const token = makeToken({
    surface: 'んですけど',
    headword: 'ん',
    reading: 'ンデスケド',
    pos1: '名詞|助動詞|助詞',
    pos2: '非自立',
  });

  assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true);
});

test('shouldExcludeTokenFromSubtitleAnnotations excludes ja-nai explanatory endings', () => {
  const tokens = [
    makeToken({
      surface: 'じゃない',
      headword: 'じゃない',
      reading: 'ジャナイ',
      partOfSpeech: PartOfSpeech.i_adjective,
      pos1: '接続詞|形容詞',
      pos2: '*|自立',
    }),
    makeToken({
      surface: 'じゃないですか',
      headword: 'じゃない',
      reading: 'ジャナイデスカ',
      partOfSpeech: PartOfSpeech.i_adjective,
      pos1: '接続詞|形容詞|助動詞|助詞',
      pos2: '*|自立|*|副助詞／並立助詞／終助詞',
    }),
  ];

  for (const token of tokens) {
    assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true, token.surface);
  }
});

test('shouldExcludeTokenFromSubtitleAnnotations excludes standalone polite copula suffix endings without POS tags', () => {
  const tokens = [
    makeToken({
      surface: 'ですよ',
      headword: 'です',
      reading: 'デスヨ',
      partOfSpeech: PartOfSpeech.other,
      pos1: '',
      pos2: '',
    }),
  ];

  for (const token of tokens) {
    assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true, token.surface);
  }
});

test('shouldExcludeTokenFromSubtitleAnnotations excludes grammar-ending patterns without enumerating variants', () => {
  const tokens = [
    makeToken({
      surface: 'ですわ',
      headword: 'です',
      reading: 'デスワ',
      partOfSpeech: PartOfSpeech.other,
      pos1: '',
      pos2: '',
    }),
    makeToken({
      surface: 'ではないですか',
      headword: 'ない',
      reading: 'デハナイデスカ',
      partOfSpeech: PartOfSpeech.other,
      pos1: '',
      pos2: '',
    }),
  ];

  for (const token of tokens) {
    assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true, token.surface);
  }
});

test('shouldExcludeTokenFromSubtitleAnnotations excludes auxiliary-stem そうだ grammar tails', () => {
  const token = makeToken({
    surface: 'そうだ',
    headword: 'そうだ',
    reading: 'ソウダ',
    pos1: '名詞|助動詞',
    pos2: '特殊',
    pos3: '助動詞語幹',
  });

  assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true);
});

test('shouldExcludeTokenFromSubtitleAnnotations keeps lexical tokens outside explanatory ending family', () => {
  const token = makeToken({
    surface: '問題',
    headword: '問題',
    reading: 'モンダイ',
    partOfSpeech: PartOfSpeech.noun,
    pos1: '名詞',
    pos2: '一般',
  });

  assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), false);
});

test('shouldExcludeTokenFromSubtitleAnnotations still excludes lexical non-independent kanji nouns from non-known annotations', () => {
  const token = makeToken({
    surface: '以外',
    headword: '以外',
    reading: 'イガイ',
    partOfSpeech: PartOfSpeech.noun,
    pos1: '名詞',
    pos2: '非自立',
    pos3: '副詞可能',
  });

  assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true);
  assert.equal(shouldExcludeTokenFromVocabularyPersistence(token), true);
});

test('shouldExcludeTokenFromSubtitleAnnotations excludes standalone particles auxiliaries and adnominals', () => {
  const tokens = [
    makeToken({
      surface: 'は',
      headword: 'は',
      reading: 'ハ',
      partOfSpeech: PartOfSpeech.particle,
      pos1: '助詞',
    }),
    makeToken({
      surface: 'です',
      headword: 'です',
      reading: 'デス',
      partOfSpeech: PartOfSpeech.bound_auxiliary,
      pos1: '助動詞',
    }),
    makeToken({
      surface: 'この',
      headword: 'この',
      reading: 'コノ',
      partOfSpeech: PartOfSpeech.other,
      pos1: '連体詞',
    }),
  ];

  for (const token of tokens) {
    assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true, token.surface);
  }
});

test('shouldExcludeTokenFromSubtitleAnnotations keeps mixed content tokens with trailing helpers', () => {
  const token = makeToken({
    surface: '行きます',
    headword: '行く',
    reading: 'イキマス',
    partOfSpeech: PartOfSpeech.verb,
    pos1: '動詞|助動詞',
    pos2: '自立',
  });

  assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), false);
});

test('shouldExcludeTokenFromSubtitleAnnotations excludes merged lexical tokens with trailing quote particles', () => {
  const token = makeToken({
    surface: 'どうしてもって',
    headword: 'どうしても',
    reading: 'ドウシテモッテ',
    partOfSpeech: PartOfSpeech.other,
    pos1: '副詞|助詞',
    pos2: '一般|格助詞',
  });

  assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true);
});

test('shouldExcludeTokenFromSubtitleAnnotations excludes kana-only demonstrative helper merges', () => {
  const token = makeToken({
    surface: 'これで',
    headword: 'これ',
    reading: 'コレデ',
    partOfSpeech: PartOfSpeech.noun,
    pos1: '名詞|助詞',
    pos2: '代名詞|格助詞',
  });

  assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true);
});

test('shouldExcludeTokenFromSubtitleAnnotations excludes kana-only non-independent noun helper merges', () => {
  const token = makeToken({
    surface: 'ことに',
    headword: '事',
    reading: 'コトニ',
    partOfSpeech: PartOfSpeech.noun,
    pos1: '名詞|助詞',
    pos2: '非自立|格助詞',
  });

  assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true);
});

test('shouldExcludeTokenFromVocabularyPersistence mirrors subtitle annotation grammar filters', () => {
  const tokens = [
    makeToken({
      surface: 'どうしてもって',
      headword: 'どうしても',
      reading: 'ドウシテモッテ',
      partOfSpeech: PartOfSpeech.other,
      pos1: '副詞|助詞',
      pos2: '一般|格助詞',
    }),
    makeToken({
      surface: 'そうだ',
      headword: 'そう',
      reading: 'ソウダ',
      partOfSpeech: PartOfSpeech.noun,
      pos1: '名詞|助動詞',
      pos2: '一般|',
      pos3: '助動詞語幹|',
    }),
  ];

  for (const token of tokens) {
    assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true, token.surface);
    assert.equal(shouldExcludeTokenFromVocabularyPersistence(token), true, token.surface);
  }
});

test('shouldExcludeTokenFromVocabularyPersistence excludes common frequency stop terms', () => {
  const tokens = [
    makeToken({
      surface: 'じゃない',
      headword: 'じゃない',
      reading: '',
      partOfSpeech: PartOfSpeech.i_adjective,
      pos1: '形容詞',
      pos2: '*|自立',
      pos3: '*',
    }),
    makeToken({
      surface: 'である',
      headword: 'である',
      reading: '',
      partOfSpeech: PartOfSpeech.verb,
      pos1: '動詞',
      pos2: '*',
      pos3: '*',
    }),
    makeToken({
      surface: '何か',
      headword: '何か',
      reading: 'なにか',
      partOfSpeech: PartOfSpeech.other,
      pos1: '名詞|助詞',
      pos2: '代名詞|副助詞／並立助詞／終助詞',
      pos3: '一般|*',
    }),
    makeToken({
      surface: '確かに',
      headword: '確かに',
      reading: 'たしかに',
      partOfSpeech: PartOfSpeech.other,
      pos1: '名詞|助詞',
      pos2: '形容動詞語幹|副詞化',
      pos3: '*',
    }),
    makeToken({
      surface: 'あなた',
      headword: '貴方',
      reading: 'あなた',
      partOfSpeech: PartOfSpeech.noun,
      pos1: '名詞',
      pos2: '代名詞',
      pos3: '一般',
    }),
  ];

  for (const token of tokens) {
    assert.equal(shouldExcludeTokenFromVocabularyPersistence(token), true, token.surface);
  }
});

test('shouldExcludeTokenFromSubtitleAnnotations excludes standalone して grammar helper fragments', () => {
  const token = makeToken({
    surface: 'して',
    headword: 'する',
    reading: 'シテ',
    partOfSpeech: PartOfSpeech.verb,
    pos1: '動詞|助詞',
    pos2: '自立|接続助詞',
  });

  assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true);
});

test('shouldExcludeTokenFromSubtitleAnnotations excludes inflected standalone して grammar helper fragments', () => {
  const token = makeToken({
    surface: 'してる',
    headword: 'する',
    reading: 'シテル',
    partOfSpeech: PartOfSpeech.verb,
    pos1: '動詞|助動詞',
    pos2: '自立|非自立',
  });

  assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true);
});

test('shouldExcludeTokenFromSubtitleAnnotations excludes standalone particle fragments without POS tags', () => {
  const token = makeToken({
    surface: 'と',
    headword: 'と',
    reading: 'ト',
    partOfSpeech: PartOfSpeech.other,
    pos1: '',
    pos2: '',
  });

  assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true);
});

test('shouldExcludeTokenFromSubtitleAnnotations excludes standalone connective particle fragments without POS tags', () => {
  const token = makeToken({
    surface: 'たって',
    headword: 'たって',
    reading: 'タッテ',
    partOfSpeech: PartOfSpeech.other,
    pos1: '',
    pos2: '',
  });

  assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true);
});

test('shouldExcludeTokenFromSubtitleAnnotations keeps lexical verbs whose reading matches connective particles', () => {
  const token = makeToken({
    surface: '立って',
    headword: '立つ',
    reading: 'タッテ',
    partOfSpeech: PartOfSpeech.verb,
    pos1: '動詞',
    pos2: '自立',
  });

  assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), false);
});

test('shouldExcludeTokenFromSubtitleAnnotations excludes rhetorical もんか grammar particle phrases', () => {
  for (const surface of ['もんか', 'ものか']) {
    const token = makeToken({
      surface,
      headword: surface,
      reading: surface === 'もんか' ? 'モンカ' : 'モノカ',
      partOfSpeech: PartOfSpeech.noun,
      pos1: '名詞|助詞',
      pos2: '非自立|副助詞／並立助詞／終助詞',
    });

    assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true, surface);
  }
});

test('shouldExcludeTokenFromSubtitleAnnotations excludes bare くれ auxiliary fragments', () => {
  const token = makeToken({
    surface: 'くれ',
    headword: '暮れ',
    reading: 'クレ',
    partOfSpeech: PartOfSpeech.noun,
    pos1: '名詞',
    pos2: '一般',
  });

  assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true);
});

test('shouldExcludeTokenFromSubtitleAnnotations excludes aru existence verbs', () => {
  for (const token of [
    makeToken({
      surface: 'ある',
      headword: 'ある',
      reading: 'アル',
      partOfSpeech: PartOfSpeech.verb,
      pos1: '動詞',
      pos2: '自立',
    }),
    makeToken({
      surface: '有る',
      headword: '有る',
      reading: 'アル',
      partOfSpeech: PartOfSpeech.verb,
      pos1: '動詞',
      pos2: '自立',
    }),
  ]) {
    assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true, token.surface);
  }
});

test('shouldExcludeTokenFromSubtitleAnnotations excludes standalone quote particle and auxiliary grammar terms', () => {
  for (const token of [
    makeToken({
      surface: 'って',
      headword: 'って',
      reading: 'ッテ',
      partOfSpeech: PartOfSpeech.other,
      pos1: '',
      pos2: '',
    }),
    makeToken({
      surface: 'べき',
      headword: 'べき',
      reading: 'ベキ',
      partOfSpeech: PartOfSpeech.other,
      pos1: '',
      pos2: '',
    }),
  ]) {
    assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true, token.surface);
  }
});

test('shouldExcludeTokenFromSubtitleAnnotations excludes single-kana surface fragments', () => {
  for (const token of [
    makeToken({
      surface: 'ふ',
      headword: '不',
      reading: 'フ',
      partOfSpeech: PartOfSpeech.other,
      pos1: '接頭詞',
      pos2: '',
    }),
    makeToken({
      surface: 'フ',
      headword: '負',
      reading: 'フ',
      partOfSpeech: PartOfSpeech.noun,
      pos1: '名詞',
      pos2: '一般',
    }),
  ]) {
    assert.equal(shouldExcludeTokenFromSubtitleAnnotations(token), true, token.surface);
  }
});

test('stripSubtitleAnnotationMetadata keeps known hover data while clearing non-known annotation fields', () => {
  const token = makeToken({
    surface: 'は',
    headword: 'は',
    reading: 'ハ',
    partOfSpeech: PartOfSpeech.particle,
    pos1: '助詞',
    isKnown: true,
    isNPlusOneTarget: true,
    isNameMatch: true,
    jlptLevel: 'N5',
    frequencyRank: 12,
  });

  assert.deepEqual(stripSubtitleAnnotationMetadata(token), {
    ...token,
    isNPlusOneTarget: false,
    isNameMatch: false,
    jlptLevel: undefined,
    frequencyRank: undefined,
  });
});

test('stripSubtitleAnnotationMetadata clears character image metadata from excluded name matches', () => {
  const token = makeToken({
    surface: 'は',
    headword: 'は',
    reading: 'ハ',
    partOfSpeech: PartOfSpeech.particle,
    pos1: '助詞',
    isNameMatch: true,
  });
  token.characterImage = {
    src: 'data:image/png;base64,AAAA',
    alt: 'は',
  };

  assert.deepEqual(stripSubtitleAnnotationMetadata(token), {
    ...token,
    isNPlusOneTarget: false,
    isNameMatch: false,
    characterImage: undefined,
    jlptLevel: undefined,
    frequencyRank: undefined,
  });
});

test('stripSubtitleAnnotationMetadata leaves content tokens unchanged', () => {
  const token = makeToken({
    surface: '猫',
    headword: '猫',
    reading: 'ネコ',
    partOfSpeech: PartOfSpeech.noun,
    pos1: '名詞',
    isKnown: true,
    jlptLevel: 'N5',
    frequencyRank: 42,
  });

  assert.strictEqual(stripSubtitleAnnotationMetadata(token), token);
});

test('annotateTokens prioritizes name matches over n+1, frequency, and JLPT when enabled', () => {
  let jlptLookupCalls = 0;
  const tokens = [
    makeToken({
      surface: 'オリヴィア',
      reading: 'オリヴィア',
      headword: 'オリヴィア',
      isNameMatch: true,
      frequencyRank: 42,
      startPos: 0,
      endPos: 5,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      getJlptLevel: () => {
        jlptLookupCalls += 1;
        return 'N2';
      },
    }),
    {
      nameMatchEnabled: true,
      minSentenceWordsForNPlusOne: 1,
    },
  );

  assert.equal(result[0]?.isNameMatch, true);
  assert.equal(result[0]?.isNPlusOneTarget, false);
  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[0]?.jlptLevel, undefined);
  assert.equal(jlptLookupCalls, 0);
});

test('annotateTokens keeps other annotations for name matches when name highlighting is disabled', () => {
  let jlptLookupCalls = 0;
  const tokens = [
    makeToken({
      surface: '山田',
      reading: 'ヤマダ',
      headword: '山田',
      isNameMatch: true,
      frequencyRank: 42,
      startPos: 0,
      endPos: 2,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      getJlptLevel: () => {
        jlptLookupCalls += 1;
        return 'N2';
      },
    }),
    {
      nameMatchEnabled: false,
      minSentenceWordsForNPlusOne: 1,
    },
  );

  assert.equal(result[0]?.isNameMatch, true);
  assert.equal(result[0]?.isNPlusOneTarget, true);
  assert.equal(result[0]?.frequencyRank, 42);
  assert.equal(result[0]?.jlptLevel, 'N2');
  assert.equal(jlptLookupCalls, 1);
});

test('annotateTokens N+1 handoff marks expected target when threshold is satisfied', () => {
  const tokens = [
    makeToken({ surface: '私', headword: '私', startPos: 0, endPos: 1 }),
    makeToken({ surface: '猫', headword: '猫', startPos: 1, endPos: 2 }),
    makeToken({
      surface: '見る',
      headword: '見る',
      partOfSpeech: PartOfSpeech.verb,
      startPos: 2,
      endPos: 4,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === '私' || text === '見る',
    }),
    { minSentenceWordsForNPlusOne: 3 },
  );

  assert.equal(result[0]?.isNPlusOneTarget, false);
  assert.equal(result[1]?.isNPlusOneTarget, true);
  assert.equal(result[2]?.isNPlusOneTarget, false);
});

test('annotateTokens does not mark kana-only unknown target as N+1', () => {
  const tokens = [
    makeToken({
      surface: '何やら',
      headword: '何やら',
      reading: 'ナニヤラ',
      pos1: '副詞',
      startPos: 0,
      endPos: 3,
    }),
    makeToken({
      surface: 'ボタン',
      headword: 'ボタン',
      reading: 'ボタン',
      pos1: '名詞',
      startPos: 3,
      endPos: 6,
    }),
    makeToken({
      surface: 'すいっち',
      headword: 'すいっち',
      reading: 'スイッチ',
      pos1: '名詞',
      startPos: 6,
      endPos: 10,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === '何やら' || text === 'ボタン',
    }),
    { minSentenceWordsForNPlusOne: 3 },
  );

  assert.equal(result[2]?.isNPlusOneTarget, false);
});

test('annotateTokens still marks kanji unknown target in otherwise eligible sentence as N+1', () => {
  const tokens = [
    makeToken({ surface: '私', headword: '私', pos1: '名詞', startPos: 0, endPos: 1 }),
    makeToken({ surface: '猫', headword: '猫', pos1: '名詞', startPos: 1, endPos: 2 }),
    makeToken({ surface: '装置…', headword: '装置', pos1: '名詞', startPos: 2, endPos: 5 }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === '私' || text === '猫',
    }),
    { minSentenceWordsForNPlusOne: 3 },
  );

  assert.equal(result[2]?.isNPlusOneTarget, true);
});

test('annotateTokens N+1 minimum sentence words counts only eligible word tokens', () => {
  const tokens = [
    makeToken({ surface: '猫', headword: '猫', startPos: 0, endPos: 1 }),
    makeToken({
      surface: 'が',
      headword: 'が',
      partOfSpeech: PartOfSpeech.particle,
      pos1: '助詞',
      startPos: 1,
      endPos: 2,
    }),
    makeToken({
      surface: 'です',
      headword: 'です',
      partOfSpeech: PartOfSpeech.bound_auxiliary,
      pos1: '助動詞',
      startPos: 2,
      endPos: 4,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === 'が' || text === 'です',
    }),
    { minSentenceWordsForNPlusOne: 3 },
  );

  assert.equal(result[0]?.isKnown, false);
  assert.equal(result[1]?.isKnown, true);
  assert.equal(result[2]?.isKnown, true);
  assert.equal(result[0]?.isNPlusOneTarget, false);
});

test('annotateTokens N+1 minimum sentence words excludes unknown tokens filtered from N+1 targeting', () => {
  const tokens = [
    makeToken({ surface: '私', headword: '私', pos1: '名詞', startPos: 0, endPos: 1 }),
    makeToken({ surface: '猫', headword: '猫', pos1: '名詞', startPos: 1, endPos: 2 }),
    makeToken({
      surface: 'スイッチ',
      headword: 'スイッチ',
      reading: 'スイッチ',
      pos1: '名詞',
      startPos: 2,
      endPos: 6,
    }),
    makeToken({ surface: '装置', headword: '装置', pos1: '名詞', startPos: 6, endPos: 8 }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === '私' || text === '猫',
    }),
    { minSentenceWordsForNPlusOne: 4 },
  );

  assert.equal(result[3]?.isNPlusOneTarget, false);
});

test('annotateTokens N+1 sentence word count respects source punctuation gaps omitted by Yomitan', () => {
  const tokens = [
    makeToken({
      surface: '私',
      headword: '私',
      pos1: '名詞',
      startPos: 0,
      endPos: 1,
    }),
    makeToken({
      surface: '猫',
      headword: '猫',
      pos1: '名詞',
      startPos: 1,
      endPos: 2,
    }),
    makeToken({
      surface: '犬',
      headword: '犬',
      pos1: '名詞',
      startPos: 2,
      endPos: 3,
    }),
    makeToken({
      surface: 'ふざけん',
      headword: 'ふざける',
      partOfSpeech: PartOfSpeech.verb,
      pos1: '動詞',
      pos2: '自立',
      startPos: 4,
      endPos: 8,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === '私' || text === '猫' || text === '犬',
    }),
    {
      minSentenceWordsForNPlusOne: 3,
      sourceText: '私猫犬！ふざけんなよ！',
    },
  );

  assert.equal(result[0]?.isNPlusOneTarget, false);
  assert.equal(result[1]?.isNPlusOneTarget, false);
  assert.equal(result[2]?.isNPlusOneTarget, false);
  assert.equal(result[3]?.isNPlusOneTarget, false);
});

test('annotateTokens N+1 sentence word count normalizes line breaks before gap detection', () => {
  const tokens = [
    makeToken({
      surface: '私',
      headword: '私',
      pos1: '名詞',
      startPos: 0,
      endPos: 1,
    }),
    makeToken({
      surface: '猫',
      headword: '猫',
      pos1: '名詞',
      startPos: 2,
      endPos: 3,
    }),
    makeToken({
      surface: '犬',
      headword: '犬',
      pos1: '名詞',
      startPos: 3,
      endPos: 4,
    }),
    makeToken({
      surface: 'ふざけん',
      headword: 'ふざける',
      partOfSpeech: PartOfSpeech.verb,
      pos1: '動詞',
      pos2: '自立',
      startPos: 5,
      endPos: 9,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === '私' || text === '猫' || text === '犬',
    }),
    {
      minSentenceWordsForNPlusOne: 3,
      sourceText: '私\r\n猫犬！ふざけんなよ！',
    },
  );

  assert.equal(result[0]?.isNPlusOneTarget, false);
  assert.equal(result[1]?.isNPlusOneTarget, false);
  assert.equal(result[2]?.isNPlusOneTarget, false);
  assert.equal(result[3]?.isNPlusOneTarget, false);
});

test('annotateTokens applies configured pos1 exclusions to both frequency and N+1', () => {
  const tokens = [
    makeToken({
      surface: '猫',
      headword: '猫',
      pos1: '名詞',
      frequencyRank: 21,
      startPos: 0,
      endPos: 1,
    }),
    makeToken({
      surface: '走る',
      headword: '走る',
      pos1: '動詞',
      partOfSpeech: PartOfSpeech.verb,
      startPos: 1,
      endPos: 3,
      frequencyRank: 22,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === '走る',
    }),
    {
      minSentenceWordsForNPlusOne: 1,
      pos1Exclusions: new Set(['名詞']),
    },
  );

  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[1]?.frequencyRank, 22);
  assert.equal(result[0]?.isNPlusOneTarget, false);
  assert.equal(result[1]?.isNPlusOneTarget, false);
});

test('annotateTokens allows previously default-excluded pos1 when removed from effective set', () => {
  const tokens = [
    makeToken({
      surface: 'まで',
      headword: 'まで',
      partOfSpeech: PartOfSpeech.other,
      pos1: '助詞',
      startPos: 0,
      endPos: 2,
      frequencyRank: 8,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps(), {
    minSentenceWordsForNPlusOne: 1,
    pos1Exclusions: new Set(),
  });

  assert.equal(result[0]?.frequencyRank, 8);
  assert.equal(result[0]?.isNPlusOneTarget, false);
});

test('annotateTokens excludes default non-independent pos2 from frequency and N+1', () => {
  const tokens = [
    makeToken({
      surface: 'になれば',
      headword: 'なる',
      partOfSpeech: PartOfSpeech.verb,
      pos1: '動詞',
      pos2: '非自立',
      startPos: 0,
      endPos: 4,
      frequencyRank: 7,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps(), {
    minSentenceWordsForNPlusOne: 1,
  });

  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[0]?.isNPlusOneTarget, false);
});

test('annotateTokens keeps known-word status for non-independent kanji noun tokens', () => {
  const tokens = [
    makeToken({
      surface: '点',
      reading: 'てん',
      headword: '点',
      partOfSpeech: PartOfSpeech.other,
      pos1: '名詞',
      pos2: '非自立',
      pos3: '一般',
      startPos: 2,
      endPos: 3,
      frequencyRank: 1384,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === '点' || text === 'てん',
      getJlptLevel: (text) => (text === '点' ? 'N3' : null),
    }),
    { minSentenceWordsForNPlusOne: 1 },
  );

  assert.equal(result[0]?.isKnown, true);
  assert.equal(result[0]?.isNPlusOneTarget, false);
  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[0]?.jlptLevel, undefined);
});

test('annotateTokens keeps known-word status for lexical non-independent kanji nouns', () => {
  const tokens = [
    makeToken({
      surface: '以外',
      reading: 'イガイ',
      headword: '以外',
      partOfSpeech: PartOfSpeech.noun,
      pos1: '名詞',
      pos2: '非自立',
      pos3: '副詞可能',
      startPos: 2,
      endPos: 4,
      frequencyRank: 437,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === '以外',
    }),
    { minSentenceWordsForNPlusOne: 1 },
  );

  assert.equal(result[0]?.isKnown, true);
  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[0]?.isNPlusOneTarget, false);
});

test('annotateTokens clears all annotations for non-independent kanji noun tokens under unified gate', () => {
  const tokens = [
    makeToken({
      surface: '者',
      reading: 'もの',
      headword: '者',
      partOfSpeech: PartOfSpeech.other,
      pos1: '名詞',
      pos2: '非自立',
      pos3: '一般',
      startPos: 0,
      endPos: 1,
      frequencyRank: 475,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps(), {
    minSentenceWordsForNPlusOne: 1,
  });

  assert.equal(result[0]?.isKnown, false);
  assert.equal(result[0]?.isNPlusOneTarget, false);
  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[0]?.jlptLevel, undefined);
});

test('annotateTokens excludes likely kana SFX tokens from frequency when POS tags are missing', () => {
  const tokens = [
    makeToken({
      surface: 'ぐわっ',
      reading: 'ぐわっ',
      headword: 'ぐわっ',
      pos1: '',
      pos2: '',
      frequencyRank: 12,
      startPos: 0,
      endPos: 3,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps(), {
    minSentenceWordsForNPlusOne: 1,
  });

  assert.equal(result[0]?.frequencyRank, undefined);
});

test('annotateTokens clears all annotations from single hiragana and katakana surface fragments', () => {
  const tokens = [
    makeToken({
      surface: 'た',
      reading: 'た',
      headword: 'た',
      pos1: '',
      pos2: '',
      partOfSpeech: PartOfSpeech.other,
      frequencyRank: 21,
      startPos: 0,
      endPos: 1,
    }),
    makeToken({
      surface: 'フ',
      reading: 'フ',
      headword: '負',
      pos1: '名詞',
      pos2: '',
      partOfSpeech: PartOfSpeech.noun,
      frequencyRank: 22,
      startPos: 1,
      endPos: 2,
    }),
    makeToken({
      surface: '山',
      reading: 'やま',
      headword: '山',
      pos1: '',
      pos2: '',
      partOfSpeech: PartOfSpeech.other,
      frequencyRank: 23,
      startPos: 2,
      endPos: 3,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps(), {
    minSentenceWordsForNPlusOne: 1,
  });

  assert.equal(result[0]?.isKnown, false);
  assert.equal(result[0]?.isNPlusOneTarget, false);
  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[0]?.jlptLevel, undefined);
  assert.equal(result[1]?.isKnown, false);
  assert.equal(result[1]?.isNPlusOneTarget, false);
  assert.equal(result[1]?.frequencyRank, undefined);
  assert.equal(result[1]?.jlptLevel, undefined);
  assert.equal(result[2]?.frequencyRank, 23);
});

test('annotateTokens keeps frequency when mecab tags classify token as content-bearing', () => {
  const tokens = [
    makeToken({
      surface: 'ふふ',
      headword: 'ふふ',
      pos1: '動詞',
      pos2: '自立',
      frequencyRank: 3014,
      startPos: 0,
      endPos: 2,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps(), {
    minSentenceWordsForNPlusOne: 1,
  });

  assert.equal(result[0]?.frequencyRank, 3014);
});

test('annotateTokens allows previously default-excluded pos2 when removed from effective set', () => {
  const tokens = [
    makeToken({
      surface: 'になれば',
      headword: 'なる',
      partOfSpeech: PartOfSpeech.verb,
      pos1: '動詞',
      pos2: '非自立',
      startPos: 0,
      endPos: 4,
      frequencyRank: 9,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps(), {
    minSentenceWordsForNPlusOne: 1,
    pos2Exclusions: new Set(),
  });

  assert.equal(result[0]?.frequencyRank, 9);
  assert.equal(result[0]?.isNPlusOneTarget, false);
});

test('annotateTokens excludes kana-only composite function/content tokens from frequency and N+1', () => {
  const tokens = [
    makeToken({
      surface: 'になれば',
      headword: 'なる',
      pos1: '助詞|動詞',
      pos2: '格助詞|自立|接続助詞',
      startPos: 0,
      endPos: 4,
      frequencyRank: 5,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps(), {
    minSentenceWordsForNPlusOne: 1,
  });

  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[0]?.isNPlusOneTarget, false);
});

test('annotateTokens excludes composite tokens when all component pos tags are excluded', () => {
  const tokens = [
    makeToken({
      surface: 'けど',
      headword: 'けど',
      pos1: '助詞|助詞',
      pos2: '接続助詞|終助詞',
      startPos: 0,
      endPos: 2,
      frequencyRank: 6,
    }),
  ];

  const result = annotateTokens(tokens, makeDeps(), {
    minSentenceWordsForNPlusOne: 1,
  });

  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[0]?.isNPlusOneTarget, false);
});

test('annotateTokens lets known words bypass the shared exclusion gate for known status only', () => {
  const tokens = [
    makeToken({
      surface: 'これで',
      headword: 'これ',
      reading: 'コレデ',
      partOfSpeech: PartOfSpeech.noun,
      pos1: '名詞|助詞',
      pos2: '代名詞|格助詞',
      startPos: 0,
      endPos: 3,
      frequencyRank: 9,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === 'これ',
      getJlptLevel: (text) => (text === 'これ' ? 'N5' : null),
    }),
    { minSentenceWordsForNPlusOne: 1 },
  );

  assert.equal(result[0]?.isKnown, true);
  assert.equal(result[0]?.isNPlusOneTarget, false);
  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[0]?.jlptLevel, undefined);
});

test('annotateTokens keeps known status while clearing other annotations for kana-only non-independent noun helper merges', () => {
  const tokens = [
    makeToken({
      surface: 'ことに',
      headword: '事',
      reading: 'コトニ',
      partOfSpeech: PartOfSpeech.noun,
      pos1: '名詞|助詞',
      pos2: '非自立|格助詞',
      startPos: 0,
      endPos: 3,
      frequencyRank: 81,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === '事',
      getJlptLevel: (text) => (text === '事' ? 'N4' : null),
    }),
    { minSentenceWordsForNPlusOne: 1 },
  );

  assert.equal(result[0]?.isKnown, true);
  assert.equal(result[0]?.isNPlusOneTarget, false);
  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[0]?.jlptLevel, undefined);
});

test('annotateTokens keeps known status while clearing other annotations for standalone auxiliary inflection fragments', () => {
  const tokens = [
    makeToken({
      surface: 'れる',
      headword: 'れる',
      reading: 'レル',
      partOfSpeech: PartOfSpeech.verb,
      pos1: '動詞',
      pos2: '接尾',
      startPos: 0,
      endPos: 2,
      frequencyRank: 18,
    }),
    makeToken({
      surface: 'れた',
      headword: 'れる',
      reading: 'レタ',
      partOfSpeech: PartOfSpeech.verb,
      pos1: '動詞|助動詞',
      pos2: '接尾|*',
      startPos: 2,
      endPos: 4,
      frequencyRank: 19,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === 'れる',
      getJlptLevel: (text) => (text === 'れる' ? 'N4' : null),
    }),
    { minSentenceWordsForNPlusOne: 1 },
  );

  for (const token of result) {
    assert.equal(token.isKnown, true, token.surface);
    assert.equal(token.isNPlusOneTarget, false, token.surface);
    assert.equal(token.frequencyRank, undefined, token.surface);
    assert.equal(token.jlptLevel, undefined, token.surface);
  }
});

test('annotateTokens keeps known status while clearing other annotations for auxiliary-only te-kureru helper spans', () => {
  const tokens = [
    makeToken({
      surface: 'てく',
      headword: 'てく',
      reading: 'テク',
      partOfSpeech: PartOfSpeech.verb,
      pos1: '助詞|動詞',
      pos2: '接続助詞|非自立',
      startPos: 0,
      endPos: 2,
      frequencyRank: 140,
    }),
    makeToken({
      surface: 'れた',
      headword: 'れる',
      reading: 'レタ',
      partOfSpeech: PartOfSpeech.verb,
      pos1: '動詞|助動詞',
      pos2: '接尾|*',
      startPos: 2,
      endPos: 4,
      frequencyRank: 19,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === 'てく' || text === 'れる',
      getJlptLevel: (text) => (text === 'てく' || text === 'れる' ? 'N4' : null),
    }),
    { minSentenceWordsForNPlusOne: 1 },
  );

  for (const token of result) {
    assert.equal(token.isKnown, true, token.surface);
    assert.equal(token.isNPlusOneTarget, false, token.surface);
    assert.equal(token.frequencyRank, undefined, token.surface);
    assert.equal(token.jlptLevel, undefined, token.surface);
  }
});

test('annotateTokens keeps lexical くれる forms eligible for annotation', () => {
  const tokens = [
    makeToken({
      surface: 'くれ',
      headword: 'くれる',
      reading: 'クレ',
      partOfSpeech: PartOfSpeech.verb,
      pos1: '動詞',
      pos2: '自立',
      startPos: 0,
      endPos: 2,
      frequencyRank: 20,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      getJlptLevel: (text) => (text === 'くれる' ? 'N4' : null),
    }),
    { minSentenceWordsForNPlusOne: 1 },
  );

  assert.equal(result[0]?.isKnown, false);
  assert.equal(result[0]?.isNPlusOneTarget, false);
  assert.equal(result[0]?.frequencyRank, 20);
  assert.equal(result[0]?.jlptLevel, 'N4');
});

test('annotateTokens keeps known status while clearing other annotations for standalone して helper fragments', () => {
  const tokens = [
    makeToken({
      surface: 'してる',
      headword: 'する',
      reading: 'シテル',
      partOfSpeech: PartOfSpeech.verb,
      pos1: '動詞|助動詞',
      pos2: '自立|非自立',
      startPos: 0,
      endPos: 3,
      frequencyRank: 22,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === 'する',
      getJlptLevel: (text) => (text === 'する' ? 'N5' : null),
    }),
    { minSentenceWordsForNPlusOne: 1 },
  );

  assert.equal(result[0]?.isKnown, true);
  assert.equal(result[0]?.isNPlusOneTarget, false);
  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[0]?.jlptLevel, undefined);
});

test('annotateTokens keeps known status while clearing other annotations for standalone particle fragments without POS tags', () => {
  const tokens = [
    makeToken({
      surface: 'と',
      headword: 'と',
      reading: 'ト',
      partOfSpeech: PartOfSpeech.other,
      pos1: '',
      pos2: '',
      startPos: 0,
      endPos: 1,
      frequencyRank: 4,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === 'と',
      getJlptLevel: (text) => (text === 'と' ? 'N5' : null),
    }),
    { minSentenceWordsForNPlusOne: 1 },
  );

  assert.equal(result[0]?.isKnown, true);
  assert.equal(result[0]?.isNPlusOneTarget, false);
  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[0]?.jlptLevel, undefined);
});

test('annotateTokens keeps known status on standalone particles when the known-word cache contains them', () => {
  const tokens = [
    makeToken({
      surface: 'に',
      headword: 'に',
      reading: 'ニ',
      partOfSpeech: PartOfSpeech.particle,
      pos1: '助詞',
      pos2: '格助詞',
      startPos: 0,
      endPos: 1,
      frequencyRank: 2,
    }),
    makeToken({
      surface: '泉',
      headword: '泉',
      reading: 'イズミ',
      partOfSpeech: PartOfSpeech.noun,
      pos1: '名詞',
      pos2: '一般',
      startPos: 1,
      endPos: 2,
      frequencyRank: 50,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === 'に' || text === '泉',
      getJlptLevel: (text) => (text === 'に' ? 'N5' : null),
    }),
    { minSentenceWordsForNPlusOne: 1 },
  );

  assert.equal(result[0]?.isKnown, true);
  assert.equal(result[0]?.isNPlusOneTarget, false);
  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[0]?.jlptLevel, undefined);
  assert.equal(result[1]?.isKnown, true);
});

test('annotateTokens does not mark standalone connective particles as N+1', () => {
  const tokens = [
    makeToken({
      surface: '逃げる',
      headword: '逃げる',
      reading: 'ニゲル',
      partOfSpeech: PartOfSpeech.verb,
      pos1: '動詞',
      pos2: '自立',
      startPos: 0,
      endPos: 3,
    }),
    makeToken({
      surface: 'たって',
      headword: 'たって',
      reading: 'タッテ',
      partOfSpeech: PartOfSpeech.other,
      pos1: '',
      pos2: '',
      startPos: 3,
      endPos: 6,
      frequencyRank: 28,
    }),
    makeToken({
      surface: '無駄',
      headword: '無駄',
      reading: 'ムダ',
      partOfSpeech: PartOfSpeech.noun,
      pos1: '名詞',
      pos2: '形容動詞語幹',
      startPos: 6,
      endPos: 8,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === '逃げる' || text === '無駄',
      getJlptLevel: (text) => (text === 'たって' ? 'N3' : null),
    }),
    { minSentenceWordsForNPlusOne: 1 },
  );

  assert.equal(result[1]?.isKnown, false);
  assert.equal(result[1]?.isNPlusOneTarget, false);
  assert.equal(result[1]?.frequencyRank, undefined);
  assert.equal(result[1]?.jlptLevel, undefined);
});

test('annotateTokens keeps known status while clearing other annotations for rhetorical もんか grammar particle phrases', () => {
  const tokens = [
    makeToken({
      surface: 'もんか',
      headword: 'もんか',
      reading: 'モンカ',
      partOfSpeech: PartOfSpeech.noun,
      pos1: '名詞|助詞',
      pos2: '非自立|副助詞／並立助詞／終助詞',
      startPos: 0,
      endPos: 3,
      frequencyRank: 69629,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === 'もんか',
      getJlptLevel: (text) => (text === 'もんか' ? 'N2' : null),
    }),
    { minSentenceWordsForNPlusOne: 1 },
  );

  assert.equal(result[0]?.isKnown, true);
  assert.equal(result[0]?.isNPlusOneTarget, false);
  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[0]?.jlptLevel, undefined);
});

test('annotateTokens keeps known status while clearing other annotations for bare くれ auxiliary fragments', () => {
  const tokens = [
    makeToken({
      surface: 'くれ',
      headword: '暮れ',
      reading: 'クレ',
      partOfSpeech: PartOfSpeech.noun,
      pos1: '名詞',
      pos2: '一般',
      startPos: 0,
      endPos: 2,
      frequencyRank: 12877,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === '暮れ',
      getJlptLevel: (text) => (text === '暮れ' ? 'N3' : null),
    }),
    { minSentenceWordsForNPlusOne: 1 },
  );

  assert.equal(result[0]?.isKnown, true);
  assert.equal(result[0]?.isNPlusOneTarget, false);
  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[0]?.jlptLevel, undefined);
});

test('annotateTokens keeps known status while clearing other annotations for aru existence verbs', () => {
  const tokens = [
    makeToken({
      surface: '有る',
      headword: '有る',
      reading: 'アル',
      partOfSpeech: PartOfSpeech.verb,
      pos1: '動詞',
      pos2: '自立',
      startPos: 0,
      endPos: 2,
      frequencyRank: 8447,
      isKnown: true,
      isNPlusOneTarget: true,
      isNameMatch: true,
      jlptLevel: 'N5',
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === '有る' || text === 'ある',
      getJlptLevel: (text) => (text === '有る' || text === 'ある' ? 'N5' : null),
    }),
    { minSentenceWordsForNPlusOne: 1 },
  );

  assert.equal(result[0]?.surface, '有る');
  assert.equal(result[0]?.headword, '有る');
  assert.equal(result[0]?.isKnown, true);
  assert.equal(result[0]?.isNPlusOneTarget, false);
  // Name matches take precedence over the annotation noise filter.
  assert.equal(result[0]?.isNameMatch, true);
  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[0]?.jlptLevel, undefined);
});

test('annotateTokens keeps known status while clearing other annotations for standalone quote particle and auxiliary grammar terms', () => {
  const tokens = [
    makeToken({
      surface: 'って',
      headword: 'って',
      reading: 'ッテ',
      partOfSpeech: PartOfSpeech.other,
      pos1: '',
      pos2: '',
      startPos: 0,
      endPos: 2,
      frequencyRank: 28,
    }),
    makeToken({
      surface: 'べき',
      headword: 'べき',
      reading: 'ベキ',
      partOfSpeech: PartOfSpeech.other,
      pos1: '',
      pos2: '',
      startPos: 2,
      endPos: 4,
      frequencyRank: 268,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === 'って' || text === 'べき',
      getJlptLevel: (text) => (text === 'って' || text === 'べき' ? 'N3' : null),
    }),
    { minSentenceWordsForNPlusOne: 1 },
  );

  for (const token of result) {
    assert.equal(token.isKnown, true, token.surface);
    assert.equal(token.isNPlusOneTarget, false, token.surface);
    assert.equal(token.frequencyRank, undefined, token.surface);
    assert.equal(token.jlptLevel, undefined, token.surface);
  }
});

test('annotateTokens keeps known status while clearing other annotations from standalone あ interjections without POS tags', () => {
  const tokens = [
    makeToken({
      surface: 'あ',
      headword: 'あ',
      reading: 'あ',
      partOfSpeech: PartOfSpeech.other,
      pos1: '',
      pos2: '',
      startPos: 0,
      endPos: 1,
      isKnown: true,
      isNPlusOneTarget: true,
      frequencyRank: 522,
      jlptLevel: 'N5',
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === 'あ',
      getJlptLevel: (text) => (text === 'あ' ? 'N5' : null),
    }),
    { minSentenceWordsForNPlusOne: 1 },
  );

  assert.equal(result[0]?.surface, 'あ');
  assert.equal(result[0]?.headword, 'あ');
  assert.equal(result[0]?.reading, 'あ');
  assert.equal(result[0]?.isKnown, true);
  assert.equal(result[0]?.isNPlusOneTarget, false);
  assert.equal(result[0]?.frequencyRank, undefined);
  assert.equal(result[0]?.jlptLevel, undefined);
});

test('annotateTokens keeps known status while clearing other annotations from expressive subtitle interjections without POS tags', () => {
  const tokens = [
    makeToken({
      surface: 'ハァ',
      headword: 'ハァ',
      reading: 'ハァ',
      partOfSpeech: PartOfSpeech.other,
      pos1: '',
      pos2: '',
      startPos: 0,
      endPos: 2,
      isKnown: true,
      isNPlusOneTarget: true,
      frequencyRank: 3007,
      jlptLevel: 'N5',
    }),
    makeToken({
      surface: 'はっ',
      headword: 'はっ',
      reading: 'ハッ',
      partOfSpeech: PartOfSpeech.other,
      pos1: '',
      pos2: '',
      startPos: 10,
      endPos: 12,
      isKnown: true,
      isNPlusOneTarget: true,
      frequencyRank: 3007,
      jlptLevel: 'N5',
    }),
    makeToken({
      surface: '猫',
      headword: '猫',
      reading: 'ネコ',
      partOfSpeech: PartOfSpeech.noun,
      pos1: '名詞',
      pos2: '一般',
      startPos: 13,
      endPos: 14,
      frequencyRank: 11,
    }),
  ];

  const result = annotateTokens(
    tokens,
    makeDeps({
      isKnownWord: (text) => text === 'ハァ' || text === 'はっ',
      getJlptLevel: (text) => (text === 'ハァ' || text === 'はっ' ? 'N5' : null),
    }),
    {
      minSentenceWordsForNPlusOne: 1,
      sourceText: 'ハァ…\n（ガーフィール）はっ！ 猫',
    },
  );

  for (const token of result.slice(0, 2)) {
    assert.equal(token.isKnown, true, token.surface);
    assert.equal(token.isNPlusOneTarget, false, token.surface);
    assert.equal(token.frequencyRank, undefined, token.surface);
    assert.equal(token.jlptLevel, undefined, token.surface);
  }
  assert.equal(result[2]?.frequencyRank, 11);
});

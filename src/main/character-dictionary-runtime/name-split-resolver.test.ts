import assert from 'node:assert/strict';
import test from 'node:test';
import { resolveJapaneseNameSplits } from './name-split-resolver';
import { splitJapaneseName, splitJapaneseNameCandidates } from './name-reading';
import { buildNameTerms } from './term-building';
import type { CharacterRecord, NameSplitToken } from './types';

function characterRecord(overrides: Partial<CharacterRecord>): CharacterRecord {
  return {
    id: 302626,
    role: 'main',
    firstNameHint: 'Shino',
    fullName: 'Shino Azuma',
    lastNameHint: 'Azuma',
    nativeName: '東紫乃',
    alternativeNames: [],
    bloodType: '',
    birthday: null,
    description: '',
    imageUrl: null,
    age: '',
    sex: '',
    voiceActors: [],
    ...overrides,
  };
}

function personNameToken(word: string, role: '姓' | '名', katakanaReading: string): NameSplitToken {
  return { word, pos1: '名詞', pos2: '固有名詞', pos3: '人名', pos4: role, katakanaReading };
}

function tokenizerFor(
  tokensByName: Record<string, NameSplitToken[]>,
): (text: string) => Promise<NameSplitToken[] | null> {
  return async (text) => tokensByName[text] ?? null;
}

test('resolveJapaneseNameSplits splits a single-kanji surname via person-name POS tags', async () => {
  const splits = await resolveJapaneseNameSplits(
    [characterRecord({})],
    tokenizerFor({
      東紫乃: [personNameToken('東', '姓', 'アズマ'), personNameToken('紫乃', '名', 'シノ')],
    }),
  );

  assert.deepEqual(splits.get('東紫乃'), { family: '東', given: '紫乃' });
});

test('resolveJapaneseNameSplits corrects a hint-length-misleading surname boundary', async () => {
  const splits = await resolveJapaneseNameSplits(
    [
      characterRecord({
        nativeName: '渡辺真奈美',
        fullName: 'Manami Watanabe',
        firstNameHint: 'Manami',
        lastNameHint: 'Watanabe',
      }),
    ],
    tokenizerFor({
      渡辺真奈美: [
        personNameToken('渡辺', '姓', 'ワタナベ'),
        personNameToken('真奈美', '名', 'マナミ'),
      ],
    }),
  );

  assert.deepEqual(splits.get('渡辺真奈美'), { family: '渡辺', given: '真奈美' });
});

test('resolveJapaneseNameSplits falls back to hint readings when POS tags are generic', async () => {
  const splits = await resolveJapaneseNameSplits(
    [
      characterRecord({
        nativeName: '鈴木みゆ',
        fullName: 'Miyu Suzuki',
        firstNameHint: 'Miyu',
        lastNameHint: 'Suzuki',
      }),
    ],
    tokenizerFor({
      鈴木みゆ: [
        personNameToken('鈴木', '姓', 'スズキ'),
        { word: 'みゆ', pos1: '名詞', pos2: '一般', pos3: '*', pos4: '*', katakanaReading: 'ミユ' },
      ],
    }),
  );

  assert.deepEqual(splits.get('鈴木みゆ'), { family: '鈴木', given: 'みゆ' });
});

test('resolveJapaneseNameSplits skips names whose tokens do not reconstruct the name', async () => {
  const splits = await resolveJapaneseNameSplits(
    [characterRecord({})],
    tokenizerFor({
      東紫乃: [personNameToken('東', '姓', 'アズマ'), personNameToken('乃', '名', 'ノ')],
    }),
  );

  assert.equal(splits.size, 0);
});

test('resolveJapaneseNameSplits skips ambiguous or untagged segmentations', async () => {
  const splits = await resolveJapaneseNameSplits(
    [
      characterRecord({
        nativeName: '担任',
        fullName: 'Tannin',
        firstNameHint: 'Tannin',
        lastNameHint: '',
      }),
    ],
    tokenizerFor({
      担任: [
        { word: '担', pos1: '名詞', pos2: '一般', pos3: '*', pos4: '*', katakanaReading: 'タン' },
        { word: '任', pos1: '名詞', pos2: '一般', pos3: '*', pos4: '*', katakanaReading: 'ニン' },
      ],
    }),
  );

  assert.equal(splits.size, 0);
});

test('resolveJapaneseNameSplits survives tokenizer failures', async () => {
  const warnings: string[] = [];
  const splits = await resolveJapaneseNameSplits(
    [characterRecord({})],
    async () => {
      throw new Error('mecab unavailable');
    },
    (message) => warnings.push(message),
  );

  assert.equal(splits.size, 0);
  assert.equal(warnings.length, 1);
  assert.match(warnings[0]!, /mecab unavailable/);
});

test('splitJapaneseName prefers a resolved split over hint-length inference', () => {
  const resolved = new Map([['東紫乃', { family: '東', given: '紫乃' }]]);

  const withResolved = splitJapaneseName('東紫乃', 'Shino', 'Azuma', resolved);
  assert.equal(withResolved.family, '東');
  assert.equal(withResolved.given, '紫乃');

  const withoutResolved = splitJapaneseName('東紫乃', 'Shino', 'Azuma');
  assert.notEqual(withoutResolved.family, '東');
});

test('splitJapaneseNameCandidates emits the runner-up boundary only for inferred splits', () => {
  const inferred = splitJapaneseNameCandidates('東紫乃', 'Shino', 'Azuma');
  assert.equal(inferred.length, 2);
  assert.deepEqual(
    inferred.map((parts) => `${parts.family}|${parts.given}`).sort(),
    ['東紫|乃', '東|紫乃'].sort(),
  );

  const resolved = new Map([['東紫乃', { family: '東', given: '紫乃' }]]);
  const trusted = splitJapaneseNameCandidates('東紫乃', 'Shino', 'Azuma', resolved);
  assert.equal(trusted.length, 1);
  assert.equal(trusted[0]!.family, '東');

  const spaced = splitJapaneseNameCandidates('須々木 心一', 'Shinichi', 'Susuki');
  assert.equal(spaced.length, 1);
});

test('buildNameTerms without resolved splits still emits both candidate surnames', () => {
  const terms = buildNameTerms(characterRecord({}));

  assert.ok(terms.includes('東'));
  assert.ok(terms.includes('紫乃'));
  assert.ok(terms.includes('東紫'));
  assert.ok(terms.includes('乃'));
});

test('buildNameTerms emits surname and given-name terms from resolved splits', () => {
  const resolved = new Map([['渡辺真奈美', { family: '渡辺', given: '真奈美' }]]);
  const terms = buildNameTerms(
    characterRecord({
      nativeName: '渡辺真奈美',
      fullName: 'Manami Watanabe',
      firstNameHint: 'Manami',
      lastNameHint: 'Watanabe',
    }),
    resolved,
  );

  assert.ok(terms.includes('渡辺'));
  assert.ok(terms.includes('真奈美'));
  assert.ok(terms.includes('渡辺真奈美'));
  assert.ok(!terms.includes('渡辺真'));
});

import assert from 'node:assert/strict';
import test from 'node:test';
import { buildNameTerms } from './term-building';
import type { CharacterRecord } from './types';

function characterRecord(overrides: Partial<CharacterRecord>): CharacterRecord {
  return {
    id: 136073,
    role: 'primary',
    firstNameHint: 'Chi-Yul',
    fullName: 'Chi-Yul Song',
    lastNameHint: 'Song',
    nativeName: '송치율',
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

test('buildNameTerms adds surname honorifics from Japanese localized aliases', () => {
  const terms = buildNameTerms(
    characterRecord({
      alternativeNames: ['Isao Mabuchi (馬渕勲)', 'Chi-Yeol (송치열)'],
    }),
  );

  assert.ok(terms.includes('馬渕勲'));
  assert.ok(terms.includes('馬渕勲さん'));
  assert.ok(terms.includes('馬渕'));
  assert.ok(terms.includes('馬渕さん'));
  assert.ok(!terms.includes('송치'));
});

test('buildNameTerms drops the disambiguator letter of a mob character name', () => {
  const terms = buildNameTerms(
    characterRecord({
      firstNameHint: '',
      lastNameHint: '',
      fullName: 'Joshi A',
      nativeName: '女子A',
    }),
  );

  // ア would match every あ〜 in the subtitles; the letter is a disambiguator
  // (Girl A / Girl B), not a name.
  assert.ok(!terms.includes('ア'));
  assert.ok(!terms.includes('アさん'));
  assert.ok(terms.includes('女子A'));
  assert.ok(terms.includes('ジョシア'));
});

test('buildNameTerms keeps a character whose whole name is one kana', () => {
  const terms = buildNameTerms(
    characterRecord({
      firstNameHint: '',
      lastNameHint: '',
      fullName: 'A',
      nativeName: 'あ',
    }),
  );

  // The mob-label rule only judges parts a name was split into; a name the
  // source gives us whole is the character's actual name.
  assert.ok(terms.includes('あ'));
  assert.ok(terms.includes('あさん'));
  // Romanized forms are never lookup targets (the subtitles are Japanese), and
  // the single-kana alias "A" transliterates to is dropped as a collision.
  assert.ok(!terms.includes('A'));
  assert.ok(!terms.includes('ア'));
});

test('buildNameTerms keeps a one-character name written in another script', () => {
  const terms = buildNameTerms(
    characterRecord({
      firstNameHint: '',
      lastNameHint: '',
      fullName: 'Byeol',
      nativeName: '별',
      alternativeNames: ['Я'],
    }),
  );

  assert.ok(terms.includes('별'));
  assert.ok(terms.includes('별さん'));
  assert.ok(terms.includes('Я'));
});

test('buildNameTerms yields nothing for a character whose only name is a bare letter', () => {
  // Documented policy rather than an oversight: a romanized name is never a
  // term on its own (the subtitles are Japanese), and the single kana a bare
  // letter transliterates to would match every あ〜 in the line.
  assert.deepEqual(
    buildNameTerms(
      characterRecord({
        firstNameHint: '',
        lastNameHint: '',
        fullName: 'A',
        nativeName: '',
      }),
    ),
    [],
  );
});

test('buildNameTerms keeps one-character split parts that are not mob labels', () => {
  const hangul = buildNameTerms(
    characterRecord({
      firstNameHint: '',
      lastNameHint: '',
      fullName: 'Byeol Kim',
      nativeName: '별 김',
    }),
  );

  assert.ok(hangul.includes('별'));
  assert.ok(hangul.includes('김'));

  const middleDot = buildNameTerms(
    characterRecord({
      firstNameHint: '',
      lastNameHint: '',
      fullName: 'A Be',
      nativeName: 'ア・ベ',
    }),
  );

  assert.ok(middleDot.includes('ア'));
  assert.ok(middleDot.includes('ベ'));
});

test('buildNameTerms keeps a single-kanji name part', () => {
  // The name is an alias, not the native name, so the parts come from the
  // space split rather than from the native-name split.
  const terms = buildNameTerms(
    characterRecord({
      firstNameHint: 'Sora',
      lastNameHint: 'Yamada',
      fullName: 'Sora Yamada',
      nativeName: '',
      alternativeNames: ['山田 空'],
    }),
  );

  assert.ok(terms.includes('山田'));
  assert.ok(terms.includes('空'));
});

test('buildNameTerms keeps a single supplementary-plane kanji name part', () => {
  // 𠮷 (U+20BB7) is a surrogate pair: a code-unit kanji check reads only the
  // high surrogate and drops the part as if it were a mob disambiguator.
  const terms = buildNameTerms(
    characterRecord({
      firstNameHint: 'Tsukasa',
      lastNameHint: 'Yoshi',
      fullName: 'Tsukasa Yoshi',
      nativeName: '',
      alternativeNames: ['𠮷 司'],
    }),
  );

  assert.ok(terms.includes('𠮷'));
  assert.ok(terms.includes('司'));
});

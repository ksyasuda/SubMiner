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

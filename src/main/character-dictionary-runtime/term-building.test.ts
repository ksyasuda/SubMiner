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
});

import test from 'node:test';
import assert from 'node:assert/strict';

import type { MergedToken } from '../types';
import { computeWordClass } from './subtitle-render.js';
import { createToken } from './subtitle-render-test-helpers.js';

test('computeWordClass preserves known and n+1 classes while adding JLPT classes', () => {
  const knownJlpt = createToken({
    isKnown: true,
    jlptLevel: 'N1',
    surface: '猫',
  });
  const nPlusOneJlpt = createToken({
    isNPlusOneTarget: true,
    jlptLevel: 'N2',
    surface: '犬',
  });

  assert.equal(computeWordClass(knownJlpt), 'word word-known word-jlpt-n1');
  assert.equal(computeWordClass(nPlusOneJlpt), 'word word-n-plus-one word-jlpt-n2');
});

test('computeWordClass applies name-match class ahead of known, n+1, frequency, and JLPT classes when enabled', () => {
  const token = createToken({
    isKnown: true,
    isNPlusOneTarget: true,
    jlptLevel: 'N2',
    frequencyRank: 10,
    surface: 'アクア',
  }) as MergedToken & { isNameMatch?: boolean };
  token.isNameMatch = true;

  assert.equal(
    computeWordClass(token, {
      nameMatchEnabled: true,
      enabled: true,
      topX: 100,
      mode: 'single',
      singleColor: '#000000',
      bandedColors: ['#000000', '#000000', '#000000', '#000000', '#000000'] as const,
    }),
    'word word-name-match',
  );
});

test('computeWordClass skips name-match class by default', () => {
  const token = createToken({
    surface: 'アクア',
  }) as MergedToken & { isNameMatch?: boolean };
  token.isNameMatch = true;

  assert.equal(computeWordClass(token), 'word');
});

test('computeWordClass skips name-match class when disabled', () => {
  const token = createToken({
    surface: 'アクア',
  }) as MergedToken & { isNameMatch?: boolean };
  token.isNameMatch = true;

  assert.equal(
    computeWordClass(token, {
      nameMatchEnabled: false,
      enabled: true,
      topX: 100,
      mode: 'single',
      singleColor: '#000000',
      bandedColors: ['#000000', '#000000', '#000000', '#000000', '#000000'] as const,
    }),
    'word',
  );
});

test('computeWordClass keeps known and N+1 color classes exclusive over frequency classes', () => {
  const known = createToken({
    isKnown: true,
    frequencyRank: 10,
    surface: '既知',
  });
  const nPlusOne = createToken({
    isNPlusOneTarget: true,
    frequencyRank: 10,
    surface: '目標',
  });
  const frequency = createToken({
    frequencyRank: 10,
    surface: '頻度',
  });

  assert.equal(
    computeWordClass(known, {
      enabled: true,
      topX: 100,
      mode: 'single',
      singleColor: '#000000',
      bandedColors: ['#000000', '#000000', '#000000', '#000000', '#000000'] as const,
    }),
    'word word-known',
  );
  assert.equal(
    computeWordClass(nPlusOne, {
      enabled: true,
      topX: 100,
      mode: 'single',
      singleColor: '#000000',
      bandedColors: ['#000000', '#000000', '#000000', '#000000', '#000000'] as const,
    }),
    'word word-n-plus-one',
  );
  assert.equal(
    computeWordClass(frequency, {
      enabled: true,
      topX: 100,
      mode: 'single',
      singleColor: '#000000',
      bandedColors: ['#000000', '#000000', '#000000', '#000000', '#000000'] as const,
    }),
    'word word-frequency-single',
  );
});

test('computeWordClass adds frequency class for single mode when rank is within topX', () => {
  const token = createToken({
    surface: '猫',
    frequencyRank: 50,
  });

  const actual = computeWordClass(token, {
    enabled: true,
    topX: 100,
    mode: 'single',
    singleColor: '#000000',
    bandedColors: ['#000000', '#000000', '#000000', '#000000', '#000000'] as const,
  });

  assert.equal(actual, 'word word-frequency-single');
});

test('computeWordClass adds frequency class when rank equals topX', () => {
  const token = createToken({
    surface: '水',
    frequencyRank: 100,
  });

  const actual = computeWordClass(token, {
    enabled: true,
    topX: 100,
    mode: 'single',
    singleColor: '#000000',
    bandedColors: ['#000000', '#000000', '#000000', '#000000', '#000000'] as const,
  });

  assert.equal(actual, 'word word-frequency-single');
});

test('computeWordClass adds frequency class for banded mode', () => {
  const token = createToken({
    surface: '犬',
    frequencyRank: 250,
  });

  const actual = computeWordClass(token, {
    enabled: true,
    topX: 1000,
    mode: 'banded',
    singleColor: '#000000',
    bandedColors: ['#111111', '#222222', '#333333', '#444444', '#555555'] as const,
  });

  assert.equal(actual, 'word word-frequency-band-2');
});

test('computeWordClass uses configured band count for banded mode', () => {
  const token = createToken({
    surface: '犬',
    frequencyRank: 2,
  });

  const actual = computeWordClass(token, {
    enabled: true,
    topX: 4,
    mode: 'banded',
    singleColor: '#000000',
    bandedColors: ['#111111', '#222222', '#333333', '#444444', '#555555'],
  } as any);

  assert.equal(actual, 'word word-frequency-band-3');
});

test('computeWordClass skips frequency class when rank is out of topX', () => {
  const token = createToken({
    surface: '犬',
    frequencyRank: 1200,
  });

  const actual = computeWordClass(token, {
    enabled: true,
    topX: 1000,
    mode: 'single',
    singleColor: '#000000',
    bandedColors: ['#000000', '#000000', '#000000', '#000000', '#000000'] as const,
  });

  assert.equal(actual, 'word');
});

test('computeWordClass adds the maturity tier class alongside word-known', () => {
  const token = createToken({
    isKnown: true,
    knownMaturity: 'mature',
    surface: '猫',
  });

  assert.equal(computeWordClass(token), 'word word-known word-maturity-mature');
});

test('computeWordClass keeps the plain known class when no maturity tier is set', () => {
  const token = createToken({
    isKnown: true,
    surface: '猫',
  });

  assert.equal(computeWordClass(token), 'word word-known');
});

test('computeWordClass composes maturity with JLPT classes', () => {
  const token = createToken({
    isKnown: true,
    knownMaturity: 'young',
    jlptLevel: 'N3',
    surface: '猫',
  });

  assert.equal(computeWordClass(token), 'word word-known word-maturity-young word-jlpt-n3');
});

test('computeWordClass gives n+1 and name matches precedence over maturity', () => {
  const nPlusOne = createToken({
    isKnown: true,
    knownMaturity: 'mature',
    isNPlusOneTarget: true,
    surface: '犬',
  });
  assert.equal(computeWordClass(nPlusOne), 'word word-n-plus-one');

  const nameMatch = createToken({
    isKnown: true,
    knownMaturity: 'mature',
    surface: 'アクア',
  }) as MergedToken & { isNameMatch?: boolean };
  nameMatch.isNameMatch = true;
  assert.equal(computeWordClass(nameMatch, { nameMatchEnabled: true }), 'word word-name-match');
});

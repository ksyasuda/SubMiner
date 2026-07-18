import test from 'node:test';
import assert from 'node:assert/strict';

import type { AnkiConnectConfig } from '../types/anki';
import {
  DEFAULT_MATURE_INTERVAL_THRESHOLD_DAYS,
  buildKnownWordMaturityTierQueries,
  classifyKnownWordNoteTier,
  getKnownWordMaturityEnabled,
  getMatureIntervalThresholdDays,
  maxKnownWordMaturityTier,
  sanitizeKnownWordMaturityTier,
} from './known-word-maturity';

function makeConfig(knownWords: AnkiConnectConfig['knownWords']): AnkiConnectConfig {
  return { url: 'http://127.0.0.1:8765', knownWords } as AnkiConnectConfig;
}

test('maturity is enabled only when both highlight and maturity flags are on', () => {
  assert.equal(
    getKnownWordMaturityEnabled(makeConfig({ highlightEnabled: true, maturityEnabled: true })),
    true,
  );
  assert.equal(
    getKnownWordMaturityEnabled(makeConfig({ highlightEnabled: false, maturityEnabled: true })),
    false,
  );
  assert.equal(
    getKnownWordMaturityEnabled(makeConfig({ highlightEnabled: true, maturityEnabled: false })),
    false,
  );
  assert.equal(getKnownWordMaturityEnabled(makeConfig({ highlightEnabled: true })), false);
  assert.equal(getKnownWordMaturityEnabled(makeConfig(undefined)), false);
});

test('mature threshold falls back to default for invalid values', () => {
  assert.equal(DEFAULT_MATURE_INTERVAL_THRESHOLD_DAYS, 21);
  assert.equal(
    getMatureIntervalThresholdDays(makeConfig({ matureThresholdDays: 30 })),
    30,
  );
  assert.equal(
    getMatureIntervalThresholdDays(makeConfig({ matureThresholdDays: 14.9 })),
    14,
  );
  assert.equal(getMatureIntervalThresholdDays(makeConfig({ matureThresholdDays: 0 })), 21);
  assert.equal(getMatureIntervalThresholdDays(makeConfig({ matureThresholdDays: -5 })), 21);
  assert.equal(
    getMatureIntervalThresholdDays(makeConfig({ matureThresholdDays: Number.NaN })),
    21,
  );
  assert.equal(getMatureIntervalThresholdDays(makeConfig({})), 21);
  assert.equal(getMatureIntervalThresholdDays(makeConfig(undefined)), 21);
});

test('tier queries append Anki search props to a deck scope query', () => {
  const queries = buildKnownWordMaturityTierQueries('deck:"Mining"', 21);
  assert.equal(queries.mature, 'deck:"Mining" prop:ivl>=21');
  assert.equal(queries.young, 'deck:"Mining" prop:ivl>=1 prop:ivl<21');
  assert.equal(queries.learning, 'deck:"Mining" is:learn');
});

test('tier queries with an empty scope query have no leading space', () => {
  const queries = buildKnownWordMaturityTierQueries('', 30);
  assert.equal(queries.mature, 'prop:ivl>=30');
  assert.equal(queries.young, 'prop:ivl>=1 prop:ivl<30');
  assert.equal(queries.learning, 'is:learn');
});

test('note classification picks the most mature matching tier', () => {
  const sets = {
    mature: new Set([1, 4]),
    young: new Set([2, 4]),
    learning: new Set([3, 4, 2]),
  };
  assert.equal(classifyKnownWordNoteTier(1, sets), 'mature');
  assert.equal(classifyKnownWordNoteTier(2, sets), 'young');
  assert.equal(classifyKnownWordNoteTier(3, sets), 'learning');
  // Note with mature, young, and learning cards: most mature card wins.
  assert.equal(classifyKnownWordNoteTier(4, sets), 'mature');
  assert.equal(classifyKnownWordNoteTier(99, sets), 'new');
});

test('maxKnownWordMaturityTier picks the higher tier and tolerates null', () => {
  assert.equal(maxKnownWordMaturityTier('mature', 'new'), 'mature');
  assert.equal(maxKnownWordMaturityTier('learning', 'young'), 'young');
  assert.equal(maxKnownWordMaturityTier('new', null), 'new');
  assert.equal(maxKnownWordMaturityTier(null, 'learning'), 'learning');
  assert.equal(maxKnownWordMaturityTier(null, null), null);
  assert.equal(maxKnownWordMaturityTier(undefined, undefined), null);
});

test('sanitizeKnownWordMaturityTier accepts only valid tiers', () => {
  assert.equal(sanitizeKnownWordMaturityTier('mature'), 'mature');
  assert.equal(sanitizeKnownWordMaturityTier('young'), 'young');
  assert.equal(sanitizeKnownWordMaturityTier('learning'), 'learning');
  assert.equal(sanitizeKnownWordMaturityTier('new'), 'new');
  assert.equal(sanitizeKnownWordMaturityTier('MATURE'), null);
  assert.equal(sanitizeKnownWordMaturityTier(''), null);
  assert.equal(sanitizeKnownWordMaturityTier(21), null);
  assert.equal(sanitizeKnownWordMaturityTier(null), null);
  assert.equal(sanitizeKnownWordMaturityTier(undefined), null);
});

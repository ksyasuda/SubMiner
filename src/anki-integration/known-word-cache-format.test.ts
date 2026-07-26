import test from 'node:test';
import assert from 'node:assert/strict';

import {
  KnownWordCacheState,
  knownWordsFromState,
  parseKnownWordCacheState,
} from './known-word-cache-format';

function parseOrThrow(value: unknown): KnownWordCacheState {
  const parsed = parseKnownWordCacheState(value);
  assert.ok(parsed, 'expected the payload to parse');
  return parsed;
}

const BASE = { refreshedAtMs: 1, scope: 'deck:test' };

test('known-word cache format reads words from every version the union covers', () => {
  assert.deepEqual(
    knownWordsFromState(parseOrThrow({ ...BASE, version: 1, words: ['する'] })),
    new Set(['する']),
  );

  assert.deepEqual(
    knownWordsFromState(
      parseOrThrow({ ...BASE, version: 2, words: ['する'], notes: { '1': ['する'] } }),
    ),
    new Set(['する']),
  );

  assert.deepEqual(
    knownWordsFromState(
      parseOrThrow({
        ...BASE,
        version: 3,
        notes: { '1': [{ word: 'する', reading: 'する' }], '2': [{ word: '猫', reading: null }] },
      }),
    ),
    new Set(['する', '猫']),
  );

  // v4 only adds `tiers`; the words a reader sees must not change.
  assert.deepEqual(
    knownWordsFromState(
      parseOrThrow({
        ...BASE,
        version: 4,
        notes: { '1': [{ word: 'する', reading: 'する' }], '2': [{ word: '猫', reading: null }] },
        tiers: { '1': 'mature', '2': 'young' },
      }),
    ),
    new Set(['する', '猫']),
  );
});

test('known-word cache format rejects payloads that are not a known cache state', () => {
  const notes = { '1': [{ word: 'する', reading: null }] };

  assert.equal(parseKnownWordCacheState(null), null);
  assert.equal(parseKnownWordCacheState('{}'), null);

  // An unknown version must not be read as an empty cache.
  assert.equal(parseKnownWordCacheState({ ...BASE, version: 99, notes }), null);

  assert.equal(
    parseKnownWordCacheState({ version: 4, scope: 'deck:test', notes, tiers: {} }),
    null,
  );
  assert.equal(parseKnownWordCacheState({ version: 4, refreshedAtMs: 1, notes, tiers: {} }), null);

  // v4 without its tiers map is a v3 payload mislabelled as v4.
  assert.equal(parseKnownWordCacheState({ ...BASE, version: 4, notes }), null);

  // v3 carries entry objects, not the bare strings v2 used.
  assert.equal(parseKnownWordCacheState({ ...BASE, version: 3, notes: { '1': ['する'] } }), null);
});

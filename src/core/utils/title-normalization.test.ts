import assert from 'node:assert/strict';
import test from 'node:test';
import { normalizeTitleIdentity } from './title-normalization';

test('normalizeTitleIdentity produces a Unicode-aware comparison key', () => {
  assert.equal(normalizeTitleIdentity('  ＢＯＣＣＨＩ・The ROCK!!  '), 'bocchi the rock');
});

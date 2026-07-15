import assert from 'node:assert/strict';
import test from 'node:test';
import * as tokenMerger from './token-merger';

test('does not expose the redundant verb non-independent predicate', () => {
  assert.equal('isVerbNonIndependent' in tokenMerger, false);
});

import test from 'node:test';
import assert from 'node:assert/strict';
import { normalizeTokenizerInput } from './normalize';

test('normalizeTokenizerInput collapses zero-width separators between Japanese segments', () => {
  const input = 'キリキリと\u200bかかってこい\nこのヘナチョコ冒険者どもめが！';
  const normalized = normalizeTokenizerInput(input);

  assert.equal(normalized, 'キリキリと かかってこい このヘナチョコ冒険者どもめが！');
});

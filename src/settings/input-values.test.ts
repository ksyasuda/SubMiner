import test from 'node:test';
import assert from 'node:assert/strict';
import { parseOptionalNumberInputValue } from './input-values';

test('parseOptionalNumberInputValue treats empty input as unset', () => {
  assert.deepEqual(parseOptionalNumberInputValue(''), { ok: true, value: undefined });
  assert.deepEqual(parseOptionalNumberInputValue('   '), { ok: true, value: undefined });
});

test('parseOptionalNumberInputValue parses finite numeric input', () => {
  assert.deepEqual(parseOptionalNumberInputValue('42'), { ok: true, value: 42 });
  assert.deepEqual(parseOptionalNumberInputValue(' 3.5 '), { ok: true, value: 3.5 });
});

test('parseOptionalNumberInputValue rejects invalid numeric input', () => {
  assert.deepEqual(parseOptionalNumberInputValue('abc'), { ok: false });
});

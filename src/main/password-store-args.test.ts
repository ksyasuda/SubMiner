import assert from 'node:assert/strict';
import test from 'node:test';
import { getPasswordStoreArg } from './password-store-args';

test('getPasswordStoreArg ignores split-form whitespace-only values', () => {
  assert.equal(getPasswordStoreArg(['SubMiner.AppImage', '--password-store', '   ']), null);
});

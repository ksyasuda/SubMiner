import test from 'node:test';
import assert from 'node:assert/strict';
import { assertSafeSshHost, shellQuote } from './ssh.js';

test('assertSafeSshHost rejects option-like hosts', () => {
  assert.throws(() => assertSafeSshHost('-oProxyCommand=touch pwned'), /looks like an option/);
  assert.throws(() => assertSafeSshHost('-lroot'), /looks like an option/);
});

test('assertSafeSshHost accepts normal destinations', () => {
  assert.doesNotThrow(() => assertSafeSshHost('macbook'));
  assert.doesNotThrow(() => assertSafeSshHost('user@192.168.1.20'));
  assert.doesNotThrow(() => assertSafeSshHost('ssh-alias'));
});

test('shellQuote escapes single quotes and wraps in quotes', () => {
  assert.equal(shellQuote('subminer'), `'subminer'`);
  assert.equal(shellQuote(`a'; rm -rf ~; '`), `'a'\\''; rm -rf ~; '\\'''`);
});

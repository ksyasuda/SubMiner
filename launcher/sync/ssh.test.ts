import test from 'node:test';
import assert from 'node:assert/strict';
import { assertSafeSshHost, runScp, shellQuote } from './ssh.js';

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

test('runScp rejects option-like local endpoints before spawning scp', () => {
  assert.throws(() => runScp('-oProxyCommand=sh', '/tmp/out.sqlite'), /looks like an option/);
  assert.throws(() => runScp('/tmp/in.sqlite', '-bad-destination'), /looks like an option/);
});

test('runScp rejects option-like remote host components', () => {
  assert.throws(
    () => runScp('-oProxyCommand=sh:/tmp/in.sqlite', '/tmp/out.sqlite'),
    /SSH host that looks like an option/,
  );
});

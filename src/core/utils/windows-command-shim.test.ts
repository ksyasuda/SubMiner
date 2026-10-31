import assert from 'node:assert/strict';
import path from 'node:path';
import test from 'node:test';

import {
  buildWindowsShimInvocation,
  isWindowsCommandShim,
  resolveWindowsCommand,
} from './windows-command-shim';

test('resolves a bare name through PATH and PATHEXT in order', () => {
  const npmDir = path.join('/fake', 'npm');
  const binDir = path.join('/fake', 'bin');
  const existing = new Set([path.join(npmDir, 'codex.CMD'), path.join(binDir, 'codex.EXE')]);
  const env = { PATH: [npmDir, binDir].join(path.delimiter), PATHEXT: '.EXE;.CMD' };

  assert.equal(
    resolveWindowsCommand('codex', env, (candidate) => existing.has(candidate)),
    path.join(npmDir, 'codex.CMD'),
  );
  assert.equal(
    resolveWindowsCommand('codex.exe', env, () => true),
    'codex.exe',
    'commands with an extension are left alone',
  );
  assert.equal(
    resolveWindowsCommand('missing', env, () => false),
    'missing',
  );
});

test('resolves an extensionless path in its own directory', () => {
  const shim = path.join('/fake', 'npm', 'codex');
  assert.equal(
    resolveWindowsCommand(shim, { PATHEXT: '.CMD' }, (candidate) => candidate === `${shim}.CMD`),
    `${shim}.CMD`,
  );
});

test('recognizes batch shims only', () => {
  assert.equal(isWindowsCommandShim('C:\\npm\\codex.cmd'), true);
  assert.equal(isWindowsCommandShim('run.BAT'), true);
  assert.equal(isWindowsCommandShim('codex.exe'), false);
});

test('quotes the shim call for cmd.exe', () => {
  const invocation = buildWindowsShimInvocation(
    'C:\\npm\\codex.cmd',
    ['-C', 'C:\\Temp dir\\a&b', '-c', 'effort="low"'],
    { ComSpec: 'C:\\Windows\\system32\\cmd.exe' },
  );

  assert.deepEqual(invocation, {
    file: 'C:\\Windows\\system32\\cmd.exe',
    args: [
      '/d',
      '/s',
      '/v:off',
      '/c',
      '"^"C:\\npm\\codex.cmd^" ^"-C^" ^"C:\\Temp^ dir\\a^&b^" ^"-c^" ^"effort=\\^"low\\^"^""',
    ],
  });
});

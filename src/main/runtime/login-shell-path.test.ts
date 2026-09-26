import test from 'node:test';
import assert from 'node:assert/strict';
import os from 'node:os';
import { applyLoginShellPath, mergePathValues, readLoginShellPath } from './login-shell-path';

const wrap = (value: string) =>
  `motd noise\n__SUBMINER_LOGIN_PATH__${value}\n__SUBMINER_LOGIN_PATH__`;

test('readLoginShellPath extracts PATH between markers, ignoring rc-file output', async () => {
  let invoked: { shell: string; args: string[] } | null = null;
  const value = await readLoginShellPath({
    env: { SHELL: '/bin/zsh' },
    runShell: async (shell, args) => {
      invoked = { shell, args };
      return wrap('/Users/me/.local/bin:/usr/bin');
    },
  });
  assert.equal(value, '/Users/me/.local/bin:/usr/bin');
  assert.equal(invoked!.shell, '/bin/zsh');
  assert.equal(invoked!.args[0], '-ilc');
});

test('readLoginShellPath returns null when markers are missing', async () => {
  const value = await readLoginShellPath({ env: {}, shell: '/bin/sh', runShell: async () => '' });
  assert.equal(value, null);
});

test('readLoginShellPath rejects relative explicit shells before execution', async () => {
  for (const shell of ['zsh', './zsh', '']) {
    await assert.rejects(
      readLoginShellPath({
        env: { SHELL: '/bin/zsh' },
        shell,
        runShell: async () => assert.fail('Invalid shell must not execute'),
      }),
      /Login shell must be an absolute path/,
    );
  }
});

test('readLoginShellPath falls back through absolute default shells', async () => {
  const originalUserInfo = os.userInfo;
  try {
    for (const shell of ['/bin/bash', 'bash', '', null]) {
      Object.defineProperty(os, 'userInfo', {
        value: () => ({ shell, username: 'test', uid: 1000, gid: 1000, homedir: '/Users/test' }),
      });
      await readLoginShellPath({
        env: { SHELL: './zsh' },
        runShell: async (selectedShell) => {
          assert.equal(selectedShell, shell === '/bin/bash' ? '/bin/bash' : '/bin/zsh');
          return wrap('/usr/bin');
        },
      });
    }
    Object.defineProperty(os, 'userInfo', {
      value: () => {
        throw new Error('User lookup failed');
      },
    });
    await readLoginShellPath({
      env: { SHELL: 'zsh' },
      runShell: async (shell) => {
        assert.equal(shell, '/bin/zsh');
        return wrap('/usr/bin');
      },
    });
  } finally {
    os.userInfo = originalUserInfo;
  }
});

test('mergePathValues puts login entries first and keeps process-only entries', () => {
  assert.equal(
    mergePathValues('/Users/me/.local/bin:/usr/bin:/bin', '/usr/bin:/bin:/usr/sbin:/sbin'),
    '/Users/me/.local/bin:/usr/bin:/bin:/usr/sbin:/sbin',
  );
});

test('applyLoginShellPath updates env.PATH so launcher detection sees shell dirs', async () => {
  const env: NodeJS.ProcessEnv = { PATH: '/usr/bin:/bin' };
  const applied = await applyLoginShellPath({
    env,
    shell: '/bin/zsh',
    runShell: async () => wrap('/Users/me/.local/bin:/usr/bin'),
  });
  assert.equal(applied, true);
  assert.equal(env.PATH, '/Users/me/.local/bin:/usr/bin:/bin');
});

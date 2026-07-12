import test from 'node:test';
import assert from 'node:assert/strict';
import { assertSafeSshHost, resolveRemoteSubminerCommand, runScp, shellQuote } from './ssh.js';

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

test('resolveRemoteSubminerCommand verifies the launcher under the remote runtime PATH', () => {
  const calls: Array<{ host: string; remoteCommand: string }> = [];
  const command = resolveRemoteSubminerCommand('macbook', null, (host, remoteCommand) => {
    calls.push({ host, remoteCommand });
    return { status: 0, stdout: '', stderr: '' };
  });

  assert.equal(
    command,
    'PATH="$HOME/.local/bin:$HOME/.bun/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:$PATH" subminer',
  );
  assert.deepEqual(calls, [
    {
      host: 'macbook',
      remoteCommand:
        'PATH="$HOME/.local/bin:$HOME/.bun/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:$PATH" subminer --help >/dev/null 2>&1',
    },
  ]);
});

test('resolveRemoteSubminerCommand falls back to the app binary in --sync-cli mode', () => {
  const probed: string[] = [];
  const command = resolveRemoteSubminerCommand('media-box', null, (_host, remoteCommand) => {
    probed.push(remoteCommand);
    return { status: remoteCommand.includes('SubMiner --sync-cli') ? 0 : 1, stdout: '', stderr: '' };
  });

  assert.match(command, / SubMiner --sync-cli$/);
  // Launcher candidates (PATH + ~/.local/bin) are tried before app binaries.
  assert.equal(probed.length, 3);
  assert.match(probed[0]!, / subminer --help /);
  assert.match(probed[1]!, / ~\/\.local\/bin\/subminer --help /);
});

test('resolveRemoteSubminerCommand probes a user override as app first, then launcher', () => {
  const probed: string[] = [];
  const asApp = resolveRemoteSubminerCommand('media-box', '/opt/SubMiner.AppImage', (_host, cmd) => {
    probed.push(cmd);
    return { status: cmd.includes('--sync-cli') ? 0 : 1, stdout: '', stderr: '' };
  });
  assert.match(asApp, /'\/opt\/SubMiner\.AppImage' --sync-cli$/);

  const asLauncher = resolveRemoteSubminerCommand('media-box', '/opt/subminer', (_host, cmd) => {
    return { status: cmd.includes('--sync-cli') ? 1 : 0, stdout: '', stderr: '' };
  });
  assert.match(asLauncher, /'\/opt\/subminer'$/);

  assert.throws(
    () =>
      resolveRemoteSubminerCommand('media-box', '/missing', () => ({
        status: 127,
        stdout: '',
        stderr: '',
      })),
    /Remote command not found on media-box: \/missing/,
  );
});

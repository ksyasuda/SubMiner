import test from 'node:test';
import assert from 'node:assert/strict';
import {
  assertSafeSshHost,
  detectRemoteShellFlavor,
  quoteForRemoteShell,
  resolveRemoteSubminerCommand,
  runScp,
  shellQuote,
  type RemoteRunResult,
} from './ssh';

function remoteResult(status: number, stdout = ''): RemoteRunResult {
  return { status, stdout, stderr: '' };
}

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
  const command = resolveRemoteSubminerCommand('macbook', null, 'posix', (host, remoteCommand) => {
    calls.push({ host, remoteCommand });
    return remoteResult(0);
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
  const command = resolveRemoteSubminerCommand('media-box', null, 'posix', (_host, remoteCommand) => {
    probed.push(remoteCommand);
    return remoteResult(remoteCommand.includes('SubMiner --sync-cli') ? 0 : 1);
  });

  assert.match(command, / SubMiner --sync-cli$/);
  // Launcher candidates (PATH + ~/.local/bin) are tried before app binaries.
  assert.equal(probed.length, 3);
  assert.match(probed[0]!, / subminer --help /);
  assert.match(probed[1]!, / ~\/\.local\/bin\/subminer --help /);
});

test('resolveRemoteSubminerCommand probes a user override as app first, then launcher', () => {
  const asApp = resolveRemoteSubminerCommand('media-box', '/opt/SubMiner.AppImage', 'posix', (_host, cmd) =>
    remoteResult(cmd.includes('--sync-cli') ? 0 : 1),
  );
  assert.match(asApp, /'\/opt\/SubMiner\.AppImage' --sync-cli$/);

  const asLauncher = resolveRemoteSubminerCommand('media-box', '/opt/subminer', 'posix', (_host, cmd) =>
    remoteResult(cmd.includes('--sync-cli') ? 1 : 0),
  );
  assert.match(asLauncher, /'\/opt\/subminer'$/);

  assert.throws(
    () => resolveRemoteSubminerCommand('media-box', '/missing', 'posix', () => remoteResult(127)),
    /Remote command not found on media-box: \/missing/,
  );
});

test('resolveRemoteSubminerCommand probes Windows install locations without a PATH prefix', () => {
  const probed: string[] = [];
  const command = resolveRemoteSubminerCommand('win-box', null, 'windows-cmd', (_host, cmd) => {
    probed.push(cmd);
    return remoteResult(cmd.includes('Programs\\SubMiner\\SubMiner.exe') ? 0 : 1);
  });

  assert.equal(command, '"%LOCALAPPDATA%\\Programs\\SubMiner\\SubMiner.exe" --sync-cli');
  assert.equal(probed[0], 'subminer --help');
  assert.equal(probed[1], '"%LOCALAPPDATA%\\SubMiner\\bin\\subminer.cmd" --help');

  const powershell = resolveRemoteSubminerCommand('win-box', null, 'windows-powershell', (_host, cmd) =>
    remoteResult(cmd.includes('Programs\\SubMiner\\SubMiner.exe') ? 0 : 1),
  );
  assert.equal(powershell, '& "$env:LOCALAPPDATA\\Programs\\SubMiner\\SubMiner.exe" --sync-cli');
});

test('resolveRemoteSubminerCommand quotes Windows overrides with double quotes', () => {
  const command = resolveRemoteSubminerCommand(
    'win-box',
    'C:/Apps/SubMiner/SubMiner.exe',
    'windows-cmd',
    (_host, cmd) => remoteResult(cmd.includes('--sync-cli') ? 0 : 1),
  );
  assert.equal(command, '"C:/Apps/SubMiner/SubMiner.exe" --sync-cli');
});

test('detectRemoteShellFlavor identifies posix, cmd, and powershell remotes', () => {
  assert.equal(
    detectRemoteShellFlavor('linux-box', (_host, cmd) =>
      cmd === 'uname -s' ? remoteResult(0, 'Linux\n') : remoteResult(1),
    ),
    'posix',
  );
  assert.equal(
    detectRemoteShellFlavor('win-box', (_host, cmd) => {
      if (cmd === 'uname -s') return remoteResult(1);
      if (cmd === 'echo %OS%') return remoteResult(0, 'Windows_NT\r\n');
      return remoteResult(1);
    }),
    'windows-cmd',
  );
  assert.equal(
    detectRemoteShellFlavor('ps-box', (_host, cmd) => {
      if (cmd === 'uname -s') return remoteResult(1);
      if (cmd === 'echo %OS%') return remoteResult(0, '%OS%\r\n');
      if (cmd === 'echo $env:OS') return remoteResult(0, 'Windows_NT\r\n');
      return remoteResult(1);
    }),
    'windows-powershell',
  );
  // Unidentifiable remotes keep the pre-detection POSIX behavior.
  assert.equal(
    detectRemoteShellFlavor('odd-box', () => remoteResult(1)),
    'posix',
  );
});

test('quoteForRemoteShell quotes per flavor and rejects unsafe Windows values', () => {
  assert.equal(quoteForRemoteShell('posix', "/tmp/it's"), `'/tmp/it'\\''s'`);
  assert.equal(
    quoteForRemoteShell('windows-cmd', 'C:/Users/First Last/AppData/Local/Temp/subminer-sync-ab'),
    '"C:/Users/First Last/AppData/Local/Temp/subminer-sync-ab"',
  );
  assert.throws(() => quoteForRemoteShell('windows-cmd', 'a"b'), /Refusing to quote/);
  assert.throws(() => quoteForRemoteShell('windows-powershell', 'a\nb'), /Refusing to quote/);
});

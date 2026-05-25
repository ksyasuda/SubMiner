import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  detectBun,
  detectLauncher,
  installLauncher,
  resolveBunInstallCommand,
  resolveLauncherInstallTarget,
  type BunSnapshot,
} from './command-line-launcher';
import { getRunCommand } from './command-line-launcher-deps';

function createBunSnapshot(status: BunSnapshot['status']): BunSnapshot {
  return {
    status,
    commandPath: status === 'ready' ? '/bin/bun' : null,
    version: status === 'ready' ? '1.3.0' : null,
    installMethod: null,
    installCommand: null,
    message: null,
  };
}

test('detectBun reports ready when bun --version succeeds on PATH', async () => {
  const snapshot = await detectBun({
    platform: 'linux',
    env: { PATH: '/usr/local/bin:/usr/bin' },
    existsSync: (candidate) => candidate === '/usr/local/bin/bun',
    accessSync: (candidate) => {
      if (candidate !== '/usr/local/bin/bun') throw new Error('not executable');
    },
    runCommand: async (command, args) => {
      assert.equal(command, '/usr/local/bin/bun');
      assert.deepEqual(args, ['--version']);
      return { exitCode: 0, stdout: '1.3.5\n', stderr: '' };
    },
  });

  assert.deepEqual(snapshot, {
    status: 'ready',
    commandPath: '/usr/local/bin/bun',
    version: '1.3.5',
    installMethod: null,
    installCommand: null,
    message: null,
  });
});

test('detectBun reports missing with an install command when bun is absent', async () => {
  const snapshot = await detectBun({
    platform: 'linux',
    env: { PATH: '/usr/bin' },
    existsSync: () => false,
    accessSync: () => {
      throw new Error('missing');
    },
    runCommand: async () => ({ exitCode: 127, stdout: '', stderr: 'missing' }),
  });

  assert.equal(snapshot.status, 'missing');
  assert.equal(snapshot.commandPath, null);
  assert.deepEqual(snapshot.installCommand, [
    'bash',
    '-lc',
    'curl -fsSL https://bun.com/install | bash',
  ]);
});

test('resolveBunInstallCommand prefers winget on Windows', () => {
  assert.deepEqual(
    resolveBunInstallCommand({
      platform: 'win32',
      env: { PATH: 'C:\\Tools' },
      existsSync: (candidate) => candidate === 'C:\\Tools\\winget.exe',
    }),
    [
      'C:\\Tools\\winget.exe',
      'install',
      '--id',
      'Oven-sh.Bun',
      '--exact',
      '--accept-package-agreements',
      '--accept-source-agreements',
    ],
  );
});

test('default runCommand preserves Windows cmd metacharacter args', async (t) => {
  if (process.platform !== 'win32') {
    t.skip('Windows cmd quoting only');
    return;
  }

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-cmd-args-'));
  const scriptPath = path.join(tempDir, 'argv.cmd');
  const outputPath = path.join(tempDir, 'argv.txt');
  t.after(() => {
    fs.rmSync(tempDir, { recursive: true, force: true });
  });
  fs.writeFileSync(
    scriptPath,
    [
      '@echo off',
      'setlocal DisableDelayedExpansion',
      '> "%SUBMINER_ARGV_OUT%" (',
      '  echo 1=%~1',
      '  echo 2=%~2',
      '  echo 3=%~3',
      '  echo 4=%~4',
      '  echo 5=%~5',
      '  echo 6=%~6',
      ')',
      '',
    ].join('\r\n'),
    'utf8',
  );

  const result = await getRunCommand({})(
    scriptPath,
    ['plain', 'has space', 'a&b', 'x|y', 'p%PATH%q', 'bang!z'],
    {
      env: { ...process.env, SUBMINER_ARGV_OUT: outputPath },
    },
  );

  assert.equal(result.exitCode, 0, result.stderr);
  assert.equal(
    fs.readFileSync(outputPath, 'utf8'),
    ['1=plain', '2=has space', '3=a&b', '4=x|y', '5=p%PATH%q', '6=bang!z', ''].join('\r\n'),
  );
});

test('resolveBunInstallCommand falls back to scoop on Windows before official installer', () => {
  assert.deepEqual(
    resolveBunInstallCommand({
      platform: 'win32',
      env: { PATH: 'C:\\Tools' },
      existsSync: (candidate) => candidate === 'C:\\Tools\\scoop.cmd',
    }),
    ['C:\\Tools\\scoop.cmd', 'install', 'bun'],
  );
});

test('resolveBunInstallCommand uses Homebrew on macOS when available', () => {
  assert.deepEqual(
    resolveBunInstallCommand({
      platform: 'darwin',
      env: { PATH: '/opt/homebrew/bin:/usr/bin' },
      existsSync: (candidate) => candidate === '/opt/homebrew/bin/brew',
      accessSync: (candidate) => {
        if (candidate !== '/opt/homebrew/bin/brew') throw new Error('not executable');
      },
    }),
    ['/opt/homebrew/bin/brew', 'install', 'bun'],
  );
});

test('detectBun reports homebrew install method from POSIX brew path', async () => {
  const snapshot = await detectBun({
    platform: 'darwin',
    env: { PATH: '/opt/homebrew/bin:/usr/bin' },
    existsSync: (candidate) => candidate === '/opt/homebrew/bin/brew',
    accessSync: (candidate) => {
      if (candidate !== '/opt/homebrew/bin/brew') throw new Error('not executable');
    },
    runCommand: async () => ({ exitCode: 127, stdout: '', stderr: 'missing' }),
  });

  assert.equal(snapshot.status, 'missing');
  assert.equal(snapshot.installMethod, 'homebrew');
});

test('resolveLauncherInstallTarget prefers writable user bin on Linux', async () => {
  const target = await resolveLauncherInstallTarget({
    platform: 'linux',
    homeDir: '/home/tester',
    env: { PATH: '/usr/bin:/home/tester/.local/bin:/tmp/bin' },
    existsSync: (candidate) =>
      candidate === '/usr/bin' ||
      candidate === '/home/tester/.local/bin' ||
      candidate === '/tmp/bin',
    accessSync: (candidate) => {
      if (candidate !== '/home/tester/.local/bin') throw new Error('not writable');
    },
  });

  assert.equal(target.status, 'not_installed');
  assert.equal(target.pathDir, '/home/tester/.local/bin');
  assert.equal(target.installPath, '/home/tester/.local/bin/subminer');
});

test('resolveLauncherInstallTarget returns not_installable without writable PATH dirs', async () => {
  const target = await resolveLauncherInstallTarget({
    platform: 'linux',
    homeDir: '/home/tester',
    env: { PATH: '/usr/bin' },
    existsSync: (candidate) => candidate === '/usr/bin',
    accessSync: () => {
      throw new Error('not writable');
    },
  });

  assert.equal(target.status, 'not_installable');
  assert.equal(target.installPath, null);
});

test('resolveLauncherInstallTarget skips Homebrew bin for empty macOS manual installs', async () => {
  const target = await resolveLauncherInstallTarget({
    platform: 'darwin',
    homeDir: '/Users/tester',
    env: { PATH: '/opt/homebrew/bin:/usr/local/bin:/Users/tester/.local/bin:/usr/bin' },
    existsSync: (candidate) =>
      candidate === '/opt/homebrew/bin' ||
      candidate === '/usr/local/bin' ||
      candidate === '/Users/tester/.local/bin' ||
      candidate === '/usr/bin',
    accessSync: (candidate) => {
      if (
        candidate !== '/opt/homebrew/bin' &&
        candidate !== '/usr/local/bin' &&
        candidate !== '/Users/tester/.local/bin'
      ) {
        throw new Error('not writable');
      }
    },
  });

  assert.equal(target.status, 'not_installed');
  assert.equal(target.pathDir, '/Users/tester/.local/bin');
  assert.equal(target.installPath, '/Users/tester/.local/bin/subminer');
});

test('resolveLauncherInstallTarget uses usr local bin for macOS manual install when user bin is absent', async () => {
  const target = await resolveLauncherInstallTarget({
    platform: 'darwin',
    homeDir: '/Users/tester',
    env: { PATH: '/opt/homebrew/bin:/usr/local/bin:/usr/bin' },
    existsSync: (candidate) =>
      candidate === '/opt/homebrew/bin' ||
      candidate === '/usr/local/bin' ||
      candidate === '/usr/bin',
    accessSync: (candidate) => {
      if (candidate !== '/opt/homebrew/bin' && candidate !== '/usr/local/bin') {
        throw new Error('not writable');
      }
    },
  });

  assert.equal(target.status, 'not_installed');
  assert.equal(target.pathDir, '/usr/local/bin');
  assert.equal(target.installPath, '/usr/local/bin/subminer');
});

test('installLauncher writes Windows cmd shim and appends user PATH once', async () => {
  const files = new Map<string, string>();
  const dirs = new Set<string>();
  const launcherResource = 'C:\\Apps\\SubMiner\\resources\\launcher\\subminer';
  files.set(launcherResource, 'launcher');
  let userPath = 'C:\\Tools;C:\\Users\\tester\\AppData\\Local\\SubMiner\\bin';
  let setPathCalls = 0;

  const snapshot = await installLauncher({
    platform: 'win32',
    localAppData: 'C:\\Users\\tester\\AppData\\Local',
    appExePath: 'C:\\Apps\\SubMiner\\SubMiner.exe',
    launcherResourcePath: launcherResource,
    env: { PATH: userPath },
    existsSync: (candidate) => files.has(candidate) || dirs.has(candidate),
    mkdirSync: (candidate) => dirs.add(candidate),
    readFileSync: (candidate) => files.get(candidate) ?? '',
    writeFileSync: (candidate, content) => files.set(candidate, content),
    getUserPath: () => userPath,
    setUserPath: (next) => {
      setPathCalls += 1;
      userPath = next;
    },
    broadcastEnvironmentChange: () => undefined,
    runCommand: async (command, args) => {
      if (command.endsWith('subminer.cmd') && args[0] === '--help') {
        return { exitCode: 0, stdout: 'help', stderr: '' };
      }
      return { exitCode: 1, stdout: '', stderr: 'unexpected' };
    },
    bunSnapshot: createBunSnapshot('ready'),
  });

  const shimPath = 'C:\\Users\\tester\\AppData\\Local\\SubMiner\\bin\\subminer.cmd';
  assert.equal(setPathCalls, 0);
  assert.match(
    files.get(shimPath) ?? '',
    /set "SUBMINER_BINARY_PATH=C:\\Apps\\SubMiner\\SubMiner\.exe"/,
  );
  assert.match(
    files.get(shimPath) ?? '',
    /bun "C:\\Apps\\SubMiner\\resources\\launcher\\subminer" %\*/,
  );
  assert.equal(snapshot.status, 'ready');
});

test('detectLauncher reports shadowed when another subminer appears earlier on PATH', async () => {
  const snapshot = await detectLauncher({
    platform: 'linux',
    homeDir: '/home/tester',
    env: { PATH: '/tmp/bin:/home/tester/.local/bin' },
    existsSync: (candidate) =>
      candidate === '/tmp/bin' ||
      candidate === '/home/tester/.local/bin' ||
      candidate === '/tmp/bin/subminer' ||
      candidate === '/home/tester/.local/bin/subminer',
    accessSync: () => undefined,
    bunSnapshot: createBunSnapshot('ready'),
  });

  assert.equal(snapshot.status, 'shadowed');
  assert.equal(snapshot.shadowedBy, '/tmp/bin/subminer');
  assert.equal(snapshot.installPath, '/home/tester/.local/bin/subminer');
});

test('detectLauncher accepts installed macOS launcher from user local bin before Homebrew target', async () => {
  const snapshot = await detectLauncher({
    platform: 'darwin',
    homeDir: '/Users/tester',
    env: { PATH: '/Users/tester/.local/bin:/opt/homebrew/bin:/usr/bin' },
    existsSync: (candidate) =>
      candidate === '/Users/tester/.local/bin' ||
      candidate === '/opt/homebrew/bin' ||
      candidate === '/Users/tester/.local/bin/subminer',
    accessSync: () => undefined,
    runCommand: async (command, args) => {
      assert.equal(command, '/Users/tester/.local/bin/subminer');
      assert.deepEqual(args, ['--help']);
      return { exitCode: 0, stdout: 'help', stderr: '' };
    },
    bunSnapshot: createBunSnapshot('ready'),
  });

  assert.equal(snapshot.status, 'ready');
  assert.equal(snapshot.commandPath, '/Users/tester/.local/bin/subminer');
  assert.equal(snapshot.installPath, '/Users/tester/.local/bin/subminer');
  assert.equal(snapshot.pathDir, '/Users/tester/.local/bin');
  assert.equal(snapshot.shadowedBy, null);
});

test('detectLauncher accepts installed macOS launcher from Homebrew bin', async () => {
  const snapshot = await detectLauncher({
    platform: 'darwin',
    homeDir: '/Users/tester',
    env: { PATH: '/opt/homebrew/bin:/usr/bin' },
    existsSync: (candidate) =>
      candidate === '/opt/homebrew/bin' || candidate === '/opt/homebrew/bin/subminer',
    accessSync: () => undefined,
    runCommand: async (command, args) => {
      assert.equal(command, '/opt/homebrew/bin/subminer');
      assert.deepEqual(args, ['--help']);
      return { exitCode: 0, stdout: 'help', stderr: '' };
    },
    bunSnapshot: createBunSnapshot('ready'),
  });

  assert.equal(snapshot.status, 'ready');
  assert.equal(snapshot.commandPath, '/opt/homebrew/bin/subminer');
  assert.equal(snapshot.installPath, '/opt/homebrew/bin/subminer');
  assert.equal(snapshot.pathDir, '/opt/homebrew/bin');
  assert.equal(snapshot.shadowedBy, null);
});

test('detectLauncher reports installed_bun_missing when launcher exists but bun is missing', async () => {
  const snapshot = await detectLauncher({
    platform: 'linux',
    homeDir: '/home/tester',
    env: { PATH: '/home/tester/.local/bin' },
    existsSync: (candidate) =>
      candidate === '/home/tester/.local/bin' || candidate === '/home/tester/.local/bin/subminer',
    accessSync: () => undefined,
    bunSnapshot: createBunSnapshot('missing'),
  });

  assert.equal(snapshot.status, 'installed_bun_missing');
});

test('detectLauncher treats stale Windows shim as not installed', async () => {
  const shimPath = 'C:\\Users\\tester\\AppData\\Local\\SubMiner\\bin\\subminer.cmd';
  const snapshot = await detectLauncher({
    platform: 'win32',
    localAppData: 'C:\\Users\\tester\\AppData\\Local',
    appExePath: 'C:\\Apps\\SubMiner\\SubMiner.exe',
    launcherResourcePath: 'C:\\Apps\\SubMiner\\resources\\launcher\\subminer',
    env: { PATH: 'C:\\Users\\tester\\AppData\\Local\\SubMiner\\bin' },
    existsSync: (candidate) =>
      candidate === shimPath || candidate === 'C:\\Users\\tester\\AppData\\Local\\SubMiner\\bin',
    readFileSync: () =>
      '@echo off\nset "SUBMINER_BINARY_PATH=C:\\Old\\SubMiner.exe"\nbun "C:\\Old\\launcher\\subminer" %*\n',
    bunSnapshot: createBunSnapshot('ready'),
  });

  assert.equal(snapshot.status, 'not_installed');
  assert.match(snapshot.message ?? '', /previous SubMiner install/);
});

test('installLauncher copies packaged launcher and chmods on POSIX', async () => {
  const files = new Map<string, string>([['/resources/launcher/subminer', 'launcher']]);
  const modes = new Map<string, number>();

  const snapshot = await installLauncher({
    platform: 'linux',
    homeDir: '/home/tester',
    env: { PATH: '/home/tester/.local/bin' },
    launcherResourcePath: '/resources/launcher/subminer',
    existsSync: (candidate) => files.has(candidate) || candidate === '/home/tester/.local/bin',
    accessSync: () => undefined,
    copyFileSync: (from, to) => files.set(to, files.get(from) ?? ''),
    chmodSync: (candidate, mode) => modes.set(candidate, mode),
    runCommand: async (command, args) => {
      assert.equal(command, '/home/tester/.local/bin/subminer');
      assert.deepEqual(args, ['--help']);
      return { exitCode: 0, stdout: 'help', stderr: '' };
    },
    bunSnapshot: createBunSnapshot('ready'),
  });

  assert.equal(files.get('/home/tester/.local/bin/subminer'), 'launcher');
  assert.equal(modes.get('/home/tester/.local/bin/subminer'), 0o755);
  assert.equal(snapshot.status, 'ready');
});

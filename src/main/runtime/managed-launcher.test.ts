import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawn, spawnSync } from 'node:child_process';
import test from 'node:test';
import {
  detectBun,
  installBun,
  installLauncher,
  refreshManagedCommandLineLauncher,
} from './command-line-launcher';
import {
  cleanupOldWindowsManagedRuntimes,
  MANAGED_LAUNCHER_MARKER,
  managedLauncherContent,
  managedLauncherPaths,
  stageManagedLauncher,
  windowsManagedRuntimePaths,
} from './managed-launcher';

function workspace(t: test.TestContext) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "subminer bundled bun's "));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  return root;
}

test('a missing packaged Bun never falls back to system Bun or runs an installer', async () => {
  let calls = 0;
  const options = {
    bundledBunPath: '/missing/private/bun',
    env: { PATH: path.dirname(process.execPath) },
    existsSync: (candidate: string) => candidate === process.execPath,
    runCommand: async () => {
      calls += 1;
      return { exitCode: 0, stdout: '1.3.5', stderr: '' };
    },
  };
  for (const snapshot of [await detectBun(options), await installBun(options)]) {
    assert.equal(snapshot.status, 'missing');
    assert.equal(snapshot.installCommand, null);
    assert.match(snapshot.message ?? '', /Reinstall SubMiner/);
  }
  assert.equal(calls, 0);
});

test('packaged POSIX launcher works without Bun on PATH and survives AppImage unmount', async (t) => {
  if (process.platform !== 'linux') return;
  const root = workspace(t);
  const resources = path.join(root, 'mounted resources');
  const bin = path.join(root, 'home', '.local', 'bin');
  fs.mkdirSync(resources, { recursive: true });
  fs.mkdirSync(bin, { recursive: true });
  const launcherResourcePath = path.join(resources, 'subminer');
  const bundledBunPath = path.join(resources, 'bun');
  fs.symlinkSync(process.execPath, bundledBunPath);
  fs.writeFileSync(
    launcherResourcePath,
    'console.log(JSON.stringify({args:process.argv.slice(2),app:process.env.SUBMINER_BINARY_PATH,managed:process.env.SUBMINER_MANAGED_LAUNCHER}));',
  );
  const appPath = path.join(root, 'SubMiner.AppImage');
  fs.writeFileSync(appPath, '#!/bin/sh\nexit 73\n', { mode: 0o755 });
  const options = {
    platform: process.platform,
    homeDir: path.join(root, 'home'),
    env: {
      HOME: path.join(root, 'home'),
      PATH: bin,
      XDG_DATA_HOME: path.join(root, 'data'),
      APPIMAGE: appPath,
    },
    appExePath: path.join(resources, 'SubMiner'),
    appVersion: '1.0.0',
    bundledBunPath,
    launcherResourcePath,
  };
  const installed = await installLauncher(options);
  assert.equal(installed.status, 'ready', installed.message ?? 'install failed');
  assert.match(fs.readFileSync(path.join(bin, 'subminer'), 'utf8'), /SubMiner managed launcher/);
  fs.renameSync(resources, path.join(root, 'unmounted'));
  const args = ['file with spaces.mkv', "single'quote", '$HOME', '$(touch nope)', '日本語'];
  const result = spawnSync(path.join(bin, 'subminer'), args, {
    env: options.env,
    encoding: 'utf8',
    timeout: 15000,
  });
  assert.equal(result.status, 0, result.stderr);
  assert.deepEqual(JSON.parse(result.stdout), { args, app: appPath, managed: '1' });
});

test('setup can install into a new user bin despite an empty GUI PATH', async (t) => {
  if (process.platform === 'win32') return;
  const root = workspace(t);
  const appPath = path.join(root, 'SubMiner.app', 'Contents', 'MacOS', 'SubMiner');
  const resources = path.join(root, 'SubMiner.app', 'Contents', 'Resources');
  fs.mkdirSync(path.dirname(appPath), { recursive: true });
  fs.mkdirSync(path.join(resources, 'bun'), { recursive: true });
  fs.mkdirSync(path.join(resources, 'launcher'), { recursive: true });
  fs.writeFileSync(appPath, '#!/bin/sh\nexit 73\n', { mode: 0o755 });
  fs.symlinkSync(process.execPath, path.join(resources, 'bun', 'bun'));
  const script = path.join(resources, 'launcher', 'subminer.js');
  fs.writeFileSync(script, 'console.log("help");');
  const snapshot = await installLauncher({
    platform: 'darwin',
    homeDir: root,
    env: { PATH: '' },
    appExePath: appPath,
    bundledBunPath: process.execPath,
    launcherResourcePath: script,
  });
  assert.equal(snapshot.status, 'not_on_path', snapshot.message ?? 'install failed');
  assert.equal(snapshot.installPath, path.join(root, '.local', 'bin', 'subminer'));
  assert.match(snapshot.message ?? '', /export PATH=/);
  const result = spawnSync(snapshot.installPath!, ['--help'], {
    env: { PATH: '' },
    encoding: 'utf8',
  });
  assert.equal(result.status, 0, result.stderr);
});

test('app upgrades refresh payloads, migrate legacy Bun launchers, and preserve custom scripts', async (t) => {
  if (process.platform !== 'linux') return;
  const root = workspace(t);
  const bin = path.join(root, 'bin');
  fs.mkdirSync(bin, { recursive: true });
  const script = path.join(root, 'resource');
  fs.writeFileSync(script, 'console.log("old");');
  const appPath = path.join(root, 'SubMiner.AppImage');
  fs.writeFileSync(appPath, '#!/bin/sh\nexit 73\n', { mode: 0o755 });
  const options = {
    platform: process.platform,
    homeDir: root,
    env: { HOME: root, PATH: bin },
    appVersion: '1',
    appExePath: appPath,
    bundledBunPath: process.execPath,
    launcherResourcePath: script,
  };
  assert.equal((await installLauncher(options)).status, 'ready');
  fs.writeFileSync(script, 'console.log("new");');
  await refreshManagedCommandLineLauncher({
    ...options,
    env: { HOME: root, PATH: '' },
    appVersion: '2',
  });
  const payload = managedLauncherPaths(options);
  assert.equal(fs.readFileSync(payload.scriptPath, 'utf8'), 'console.log("new");');
  fs.writeFileSync(path.join(bin, 'subminer'), '#!/usr/bin/env bun\n// SubMiner launcher\n');
  await refreshManagedCommandLineLauncher({ ...options, appVersion: '3' });
  assert.match(fs.readFileSync(path.join(bin, 'subminer'), 'utf8'), /SubMiner managed launcher/);
  fs.writeFileSync(path.join(bin, 'subminer'), '#!/bin/sh\necho standalone\n');
  await refreshManagedCommandLineLauncher({ ...options, appVersion: '4' });
  assert.equal(fs.readFileSync(payload.versionPath, 'utf8'), '3');
});

test('Windows wrapper discovers the configured app and its versioned private runtime', () => {
  const content = managedLauncherContent({
    platform: 'win32',
    appPath: 'C:\\Apps 100% !\\SubMiner.exe',
  });
  assert.ok(content.includes(MANAGED_LAUNCHER_MARKER));
  assert.ok(content.includes('setlocal DisableDelayedExpansion'));
  assert.ok(content.includes('set "SUBMINER_BINARY_PATH=C:\\Apps 100%% !\\SubMiner.exe"'));
  assert.ok(content.includes('%SUBMINER_RESOURCES_PATH%\\launcher\\version'));
  assert.ok(
    content.includes(
      'set "SUBMINER_BUN_PATH=%LOCALAPPDATA%\\SubMiner\\launcher-runtime\\%SUBMINER_APP_VERSION%\\bun.exe"',
    ),
  );
  assert.ok(
    content.includes('"%SUBMINER_BUN_PATH%" "%SUBMINER_RESOURCES_PATH%\\launcher\\subminer.js" %*'),
  );
  assert.ok(content.includes('exit /b %errorlevel%'));
});

test('Windows managed runtime path is absolute, versioned, and injectable', () => {
  const paths = windowsManagedRuntimePaths({
    platform: 'win32',
    localAppData: 'D:\\Profiles\\テスト User\\AppData\\Local',
    appVersion: '1.2.3-beta.4',
  });
  assert.equal(
    paths.bunPath,
    'D:\\Profiles\\テスト User\\AppData\\Local\\SubMiner\\launcher-runtime\\1.2.3-beta.4\\bun.exe',
  );
  assert.ok(path.win32.isAbsolute(paths.bunPath));
  assert.throws(
    () =>
      windowsManagedRuntimePaths({
        platform: 'win32',
        localAppData: 'relative',
        appVersion: '1.2.3',
      }),
    /must be an absolute Windows path/,
  );
  assert.throws(
    () =>
      windowsManagedRuntimePaths({
        platform: 'win32',
        localAppData: 'C:\\Users\\tester\\AppData\\Local',
        appVersion: '1:2',
      }),
    /not a valid directory name/,
  );
});

test('Windows managed launcher forwards arguments without a system Bun', async (t) => {
  if (process.platform !== 'win32') return;
  const root = workspace(t);
  const appDirectory = path.join(root, 'Installed App');
  const appPath = path.join(appDirectory, 'SubMiner.exe');
  const launcherDirectory = path.join(appDirectory, 'resources', 'launcher');
  const script = path.join(launcherDirectory, 'subminer.js');
  fs.mkdirSync(launcherDirectory, { recursive: true });
  fs.copyFileSync(process.execPath, appPath);
  fs.writeFileSync(script, 'console.log(JSON.stringify(process.argv.slice(2)));');
  fs.writeFileSync(path.join(launcherDirectory, 'version'), '1.0.0');
  const options = {
    platform: process.platform,
    env: { ...process.env, PATH: '', LOCALAPPDATA: root },
    bundledBunPath: process.execPath,
    launcherResourcePath: script,
    appExePath: appPath,
    appVersion: '1.0.0',
    localAppData: root,
    getUserPath: () => '',
    setUserPath: () => {},
    broadcastEnvironmentChange: () => {},
  };
  const snapshot = await installLauncher(options);
  assert.equal(snapshot.status, 'ready', snapshot.message ?? 'install failed');
  const { getRunCommand } = await import('./command-line-launcher-deps');
  const args = ['spaces here', 'a&b', 'p%TEMP%q', 'bang!z', 'say "hi"', '日本語'];
  const result = await getRunCommand({})(snapshot.installPath!, args, { env: options.env });
  assert.equal(result.exitCode, 0, result.stderr);
  assert.deepEqual(JSON.parse(result.stdout), args);
});

test('Windows stages a new runtime version while the prior Bun executable is running', async (t) => {
  if (process.platform !== 'win32') return;
  const root = workspace(t);
  const bundledDirectory = path.join(root, 'packaged', 'bun');
  const bundledBunPath = path.join(bundledDirectory, 'bun.exe');
  const launcherResourcePath = path.join(root, 'packaged', 'launcher', 'subminer');
  fs.mkdirSync(path.join(bundledDirectory, 'licenses'), { recursive: true });
  fs.mkdirSync(path.dirname(launcherResourcePath), { recursive: true });
  fs.copyFileSync(process.execPath, bundledBunPath);
  fs.writeFileSync(path.join(bundledDirectory, 'licenses', 'Bun-LICENSE.md'), 'license');
  fs.writeFileSync(launcherResourcePath, 'console.log("launcher");');

  const first = stageManagedLauncher({
    platform: 'win32',
    localAppData: root,
    appVersion: '1.0.0',
    bundledBunPath,
    launcherResourcePath,
  });
  const running = spawn(first.bunPath, ['-e', 'setInterval(() => {}, 1000)']);
  await new Promise<void>((resolve, reject) => {
    running.once('spawn', resolve);
    running.once('error', reject);
  });

  const expectedSecond = windowsManagedRuntimePaths({
    platform: 'win32',
    localAppData: root,
    appVersion: '2.0.0',
  });
  try {
    const second = stageManagedLauncher({
      platform: 'win32',
      localAppData: root,
      appVersion: '2.0.0',
      bundledBunPath,
      launcherResourcePath,
    });
    assert.notEqual(second.bunPath, first.bunPath);
    assert.ok(fs.existsSync(second.bunPath));
    cleanupOldWindowsManagedRuntimes({
      platform: 'win32',
      localAppData: root,
      appVersion: '2.0.0',
    });
    assert.ok(fs.existsSync(first.bunPath));
    assert.ok(fs.existsSync(path.join(path.dirname(first.bunPath), 'licenses', 'Bun-LICENSE.md')));
    assert.equal(
      fs.readFileSync(
        path.join(path.dirname(second.bunPath), 'licenses', 'Bun-LICENSE.md'),
        'utf8',
      ),
      'license',
    );
  } finally {
    if (running.exitCode === null) {
      const exited = new Promise<void>((resolve) => running.once('exit', () => resolve()));
      running.kill();
      await exited;
    }
  }
  cleanupOldWindowsManagedRuntimes({
    platform: 'win32',
    localAppData: root,
    appVersion: '2.0.0',
  });
  assert.equal(fs.existsSync(path.dirname(first.bunPath)), false);
  assert.ok(fs.existsSync(expectedSecond.bunPath));
});

test('Windows cleanup removes an obsolete runtime with no Bun executable', (t) => {
  if (process.platform !== 'win32') return;
  const root = workspace(t);
  const current = windowsManagedRuntimePaths({
    platform: 'win32',
    localAppData: root,
    appVersion: '2.0.0',
  });
  const obsolete = windowsManagedRuntimePaths({
    platform: 'win32',
    localAppData: root,
    appVersion: '1.0.0',
  });
  fs.mkdirSync(path.join(path.dirname(obsolete.bunPath), 'licenses'), { recursive: true });
  fs.writeFileSync(
    path.join(path.dirname(obsolete.bunPath), 'licenses', 'Bun-LICENSE.md'),
    'license',
  );
  fs.mkdirSync(path.dirname(current.bunPath), { recursive: true });

  cleanupOldWindowsManagedRuntimes({
    platform: 'win32',
    localAppData: root,
    appVersion: '2.0.0',
  });

  assert.equal(fs.existsSync(path.dirname(obsolete.bunPath)), false);
  assert.ok(fs.existsSync(path.dirname(current.bunPath)));
});

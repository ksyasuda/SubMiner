import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  detectInstalledFirstRunPlugin,
  detectInstalledFirstRunPluginCandidates,
  detectInstalledMpvPlugin,
  removeLegacyMpvPluginCandidates,
  resolvePackagedFirstRunPluginAssets,
  resolvePackagedRuntimePluginPath,
} from './first-run-setup-plugin';
import { resolveDefaultMpvInstallPaths } from '../../shared/setup-state';

function withTempDir(fn: (dir: string) => Promise<void> | void): Promise<void> | void {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-first-run-plugin-test-'));
  const result = fn(dir);
  if (result instanceof Promise) {
    return result.finally(() => {
      fs.rmSync(dir, { recursive: true, force: true });
    });
  }
  fs.rmSync(dir, { recursive: true, force: true });
}

test('resolvePackagedFirstRunPluginAssets finds packaged plugin assets', () => {
  withTempDir((root) => {
    const resourcesPath = path.join(root, 'resources');
    const pluginRoot = path.join(resourcesPath, 'plugin');
    fs.mkdirSync(path.join(pluginRoot, 'subminer'), { recursive: true });
    fs.writeFileSync(path.join(pluginRoot, 'subminer', 'main.lua'), '-- plugin');
    fs.writeFileSync(path.join(pluginRoot, 'subminer.conf'), 'configured=true\n');

    const resolved = resolvePackagedFirstRunPluginAssets({
      dirname: path.join(root, 'dist', 'main', 'runtime'),
      appPath: path.join(root, 'app'),
      resourcesPath,
    });

    assert.deepEqual(resolved, {
      pluginDirSource: path.join(pluginRoot, 'subminer'),
      pluginConfigSource: path.join(pluginRoot, 'subminer.conf'),
    });
  });
});

test('resolvePackagedRuntimePluginPath returns packaged plugin entrypoint', () => {
  withTempDir((root) => {
    const resourcesPath = path.join(root, 'resources');
    const pluginRoot = path.join(resourcesPath, 'plugin');
    const entrypoint = path.join(pluginRoot, 'subminer', 'main.lua');
    fs.mkdirSync(path.dirname(entrypoint), { recursive: true });
    fs.writeFileSync(entrypoint, '-- plugin');
    fs.writeFileSync(path.join(pluginRoot, 'subminer.conf'), 'configured=true\n');

    assert.equal(
      resolvePackagedRuntimePluginPath({
        dirname: path.join(root, 'dist', 'main', 'runtime'),
        appPath: path.join(root, 'app'),
        resourcesPath,
      }),
      entrypoint,
    );
  });
});

test('detectInstalledFirstRunPlugin detects plugin installed in canonical mpv config location on macOS', () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const installPaths = resolveDefaultMpvInstallPaths('darwin', homeDir);
    const pluginDir = path.join(homeDir, '.config', 'mpv', 'scripts', 'subminer');
    const pluginEntrypointPath = path.join(pluginDir, 'main.lua');

    fs.mkdirSync(pluginDir, { recursive: true });
    fs.mkdirSync(path.dirname(pluginEntrypointPath), { recursive: true });
    fs.writeFileSync(pluginEntrypointPath, '-- plugin');

    assert.equal(detectInstalledFirstRunPlugin(installPaths), true);
  });
});

test('detectInstalledFirstRunPlugin ignores scoped plugin layout path', () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const xdgConfigHome = path.join(root, 'xdg');
    const installPaths = resolveDefaultMpvInstallPaths('darwin', homeDir, xdgConfigHome);
    const pluginDir = path.join(xdgConfigHome, 'mpv', 'scripts', '@plugin', 'subminer');
    const pluginEntrypointPath = path.join(pluginDir, 'main.lua');

    fs.mkdirSync(pluginDir, { recursive: true });
    fs.mkdirSync(path.dirname(pluginEntrypointPath), { recursive: true });
    fs.writeFileSync(pluginEntrypointPath, '-- plugin');

    assert.equal(detectInstalledFirstRunPlugin(installPaths), false);
  });
});

test('detectInstalledFirstRunPlugin ignores legacy loader file', () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const installPaths = resolveDefaultMpvInstallPaths('darwin', homeDir);
    const legacyLoaderPath = path.join(installPaths.scriptsDir, 'subminer.lua');

    fs.mkdirSync(path.dirname(legacyLoaderPath), { recursive: true });
    fs.writeFileSync(legacyLoaderPath, '-- plugin');

    assert.equal(detectInstalledFirstRunPlugin(installPaths), false);
  });
});

test('detectInstalledFirstRunPluginCandidates returns all legacy autoload entries without script opts', () => {
  withTempDir((root) => {
    const homeDir = path.posix.join(root, 'home');
    const xdgConfigHome = path.posix.join(root, 'xdg');
    const installPaths = resolveDefaultMpvInstallPaths('linux', homeDir, xdgConfigHome);
    const directoryInstall = installPaths.pluginDir;
    const legacyScript = path.posix.join(installPaths.scriptsDir, 'subminer.lua');
    const legacyLoader = path.posix.join(installPaths.scriptsDir, 'subminer-loader.lua');

    fs.mkdirSync(directoryInstall, { recursive: true });
    fs.writeFileSync(path.posix.join(directoryInstall, 'main.lua'), '-- plugin');
    fs.writeFileSync(legacyScript, '-- legacy plugin');
    fs.writeFileSync(legacyLoader, '-- legacy loader');
    fs.mkdirSync(path.dirname(installPaths.pluginConfigPath), { recursive: true });
    fs.writeFileSync(installPaths.pluginConfigPath, 'socket_path=/tmp/subminer-socket\n');

    const candidates = detectInstalledFirstRunPluginCandidates({
      platform: 'linux',
      homeDir,
      xdgConfigHome,
    });

    assert.deepEqual(
      candidates.map((candidate) => candidate.path).sort(),
      [directoryInstall, legacyLoader, legacyScript].sort(),
    );
    assert.equal(
      candidates.some((candidate) => candidate.path === installPaths.pluginConfigPath),
      false,
    );
  });
});

test('detectInstalledFirstRunPluginCandidates includes Windows portable mpv scripts', () => {
  withTempDir((root) => {
    const homeDir = path.win32.join('C:\\Users', 'tester');
    const appDataDir = path.win32.join(root, 'AppData', 'Roaming');
    const mpvExecutablePath = path.win32.join(root, 'mpv', 'mpv.exe');
    const portablePluginDir = path.win32.join(
      path.win32.dirname(mpvExecutablePath),
      'portable_config',
      'scripts',
      'subminer',
    );
    const portableLegacyScript = path.win32.join(
      path.win32.dirname(mpvExecutablePath),
      'portable_config',
      'scripts',
      'subminer.lua',
    );
    const existing = new Set([portablePluginDir, portableLegacyScript]);

    const candidates = detectInstalledFirstRunPluginCandidates({
      platform: 'win32',
      homeDir,
      appDataDir,
      mpvExecutablePath,
      existsSync: (candidate) => existing.has(candidate),
    });

    assert.deepEqual(
      candidates.map((candidate) => candidate.path),
      [portablePluginDir, portableLegacyScript],
    );
  });
});

test('detectInstalledMpvPlugin prefers Windows portable plugin and parses version', () => {
  const homeDir = 'C:\\Users\\tester';
  const appDataDir = 'C:\\Users\\tester\\AppData\\Roaming';
  const mpvExecutablePath = 'C:\\tools\\mpv\\mpv.exe';
  const portableEntrypoint = 'C:\\tools\\mpv\\portable_config\\scripts\\subminer\\main.lua';
  const portableVersion = 'C:\\tools\\mpv\\portable_config\\scripts\\subminer\\version.lua';
  const appDataEntrypoint = 'C:\\Users\\tester\\AppData\\Roaming\\mpv\\scripts\\subminer\\main.lua';
  const existing = new Set([portableEntrypoint, portableVersion, appDataEntrypoint]);

  const detection = detectInstalledMpvPlugin({
    platform: 'win32',
    homeDir,
    appDataDir,
    mpvExecutablePath,
    existsSync: (candidate) => existing.has(candidate),
    readFileSync: (candidate) =>
      candidate === portableVersion ? 'return { version = "0.12.0" }' : '',
  });

  assert.equal(detection.installed, true);
  assert.equal(detection.path, portableEntrypoint);
  assert.equal(detection.version, '0.12.0');
  assert.equal(detection.source, 'portable-config');
});

test('detectInstalledMpvPlugin detects Linux legacy single-file plugin without version', () => {
  withTempDir((root) => {
    const homeDir = path.posix.join(root, 'home');
    const legacyPath = path.posix.join(homeDir, '.config', 'mpv', 'scripts', 'subminer-loader.lua');
    fs.mkdirSync(path.posix.dirname(legacyPath), { recursive: true });
    fs.writeFileSync(legacyPath, '-- legacy');

    const detection = detectInstalledMpvPlugin({
      platform: 'linux',
      homeDir,
    });

    assert.equal(detection.installed, true);
    assert.equal(detection.path, legacyPath);
    assert.equal(detection.version, null);
    assert.equal(detection.source, 'legacy-file');
  });
});

test('removeLegacyMpvPluginCandidates trashes candidates and reports partial failures', async () => {
  const calls: string[] = [];
  const result = await removeLegacyMpvPluginCandidates({
    candidates: [
      { path: '/tmp/mpv/scripts/subminer', kind: 'directory' },
      { path: '/tmp/mpv/scripts/subminer.lua', kind: 'file' },
    ],
    trashItem: async (candidate) => {
      calls.push(candidate);
      if (candidate.endsWith('subminer.lua')) {
        throw new Error('permission denied');
      }
    },
  });

  assert.deepEqual(calls, ['/tmp/mpv/scripts/subminer', '/tmp/mpv/scripts/subminer.lua']);
  assert.equal(result.ok, false);
  assert.deepEqual(result.removedPaths, ['/tmp/mpv/scripts/subminer']);
  assert.deepEqual(result.failedPaths, [
    { path: '/tmp/mpv/scripts/subminer.lua', message: 'permission denied' },
  ]);
});

test('detectInstalledFirstRunPlugin requires main.lua in subminer directory', () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const installPaths = resolveDefaultMpvInstallPaths('darwin', homeDir);
    const pluginDir = path.join(installPaths.scriptsDir, 'subminer');

    fs.mkdirSync(pluginDir, { recursive: true });
    fs.writeFileSync(path.join(pluginDir, 'not_main.lua'), '-- plugin');

    assert.equal(detectInstalledFirstRunPlugin(installPaths), false);
  });
});

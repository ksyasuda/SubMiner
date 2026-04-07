import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  detectInstalledFirstRunPlugin,
  installFirstRunPluginToDefaultLocation,
  resolvePackagedFirstRunPluginAssets,
  syncInstalledFirstRunPluginBinaryPath,
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

test('installFirstRunPluginToDefaultLocation installs plugin and backs up existing files', () => {
  withTempDir((root) => {
    const resourcesPath = path.join(root, 'resources');
    const pluginRoot = path.join(resourcesPath, 'plugin');
    const homeDir = path.join(root, 'home');
    const xdgConfigHome = path.join(root, 'xdg');
    const installPaths = resolveDefaultMpvInstallPaths('linux', homeDir, xdgConfigHome);

    fs.mkdirSync(path.join(pluginRoot, 'subminer'), { recursive: true });
    fs.writeFileSync(path.join(pluginRoot, 'subminer', 'main.lua'), '-- packaged plugin');
    fs.writeFileSync(path.join(pluginRoot, 'subminer.conf'), 'configured=true\n');

    fs.mkdirSync(path.dirname(installPaths.pluginEntrypointPath), { recursive: true });
    fs.mkdirSync(installPaths.pluginDir, { recursive: true });
    fs.mkdirSync(path.dirname(installPaths.pluginConfigPath), { recursive: true });
    fs.writeFileSync(path.join(installPaths.scriptsDir, 'subminer-loader.lua'), '-- old loader');
    fs.writeFileSync(path.join(installPaths.pluginDir, 'old.lua'), '-- old plugin');
    fs.writeFileSync(installPaths.pluginConfigPath, 'old=true\n');

    const result = installFirstRunPluginToDefaultLocation({
      platform: 'linux',
      homeDir,
      xdgConfigHome,
      dirname: path.join(root, 'dist', 'main', 'runtime'),
      appPath: path.join(root, 'app'),
      resourcesPath,
      binaryPath: '/Applications/SubMiner.app/Contents/MacOS/SubMiner',
    });

    assert.equal(result.ok, true);
    assert.equal(result.pluginInstallStatus, 'installed');
    assert.equal(detectInstalledFirstRunPlugin(installPaths), true);
    assert.equal(fs.readFileSync(installPaths.pluginEntrypointPath, 'utf8'), '-- packaged plugin');
    assert.equal(
      fs.readFileSync(installPaths.pluginConfigPath, 'utf8'),
      'configured=true\nbinary_path=/Applications/SubMiner.app/Contents/MacOS/SubMiner\n',
    );

    const scriptsDirEntries = fs.readdirSync(installPaths.scriptsDir);
    const scriptOptsEntries = fs.readdirSync(installPaths.scriptOptsDir);
    assert.equal(
      scriptsDirEntries.some((entry) => entry.startsWith('subminer.bak.')),
      true,
    );
    assert.equal(
      scriptsDirEntries.some((entry) => entry.startsWith('subminer-loader.lua.bak.')),
      true,
    );
    assert.equal(
      scriptOptsEntries.some((entry) => entry.startsWith('subminer.conf.bak.')),
      true,
    );
  });
});

test('installFirstRunPluginToDefaultLocation installs plugin to Windows mpv defaults', () => {
  if (process.platform !== 'win32') {
    return;
  }
  withTempDir((root) => {
    const resourcesPath = path.join(root, 'resources');
    const pluginRoot = path.join(resourcesPath, 'plugin');
    const homeDir = path.join(root, 'home');
    const installPaths = resolveDefaultMpvInstallPaths('win32', homeDir);

    fs.mkdirSync(path.join(pluginRoot, 'subminer'), { recursive: true });
    fs.writeFileSync(path.join(pluginRoot, 'subminer', 'main.lua'), '-- packaged plugin');
    fs.writeFileSync(path.join(pluginRoot, 'subminer.conf'), 'configured=true\n');

    const result = installFirstRunPluginToDefaultLocation({
      platform: 'win32',
      homeDir,
      dirname: path.join(root, 'dist', 'main', 'runtime'),
      appPath: path.join(root, 'app'),
      resourcesPath,
      binaryPath: 'C:\\Program Files\\SubMiner\\SubMiner.exe',
    });

    assert.equal(result.ok, true);
    assert.equal(result.pluginInstallStatus, 'installed');
    assert.equal(detectInstalledFirstRunPlugin(installPaths), true);
    assert.equal(fs.readFileSync(installPaths.pluginEntrypointPath, 'utf8'), '-- packaged plugin');
    assert.equal(
      fs.readFileSync(installPaths.pluginConfigPath, 'utf8'),
      'configured=true\nbinary_path=C:\\Program Files\\SubMiner\\SubMiner.exe\n',
    );
  });
});

test('installFirstRunPluginToDefaultLocation rewrites Windows plugin socket_path', () => {
  if (process.platform !== 'win32') {
    return;
  }
  withTempDir((root) => {
    const resourcesPath = path.join(root, 'resources');
    const pluginRoot = path.join(resourcesPath, 'plugin');
    const homeDir = path.join(root, 'home');
    const installPaths = resolveDefaultMpvInstallPaths('win32', homeDir);

    fs.mkdirSync(path.join(pluginRoot, 'subminer'), { recursive: true });
    fs.writeFileSync(path.join(pluginRoot, 'subminer', 'main.lua'), '-- packaged plugin');
    fs.writeFileSync(
      path.join(pluginRoot, 'subminer.conf'),
      'binary_path=\nsocket_path=/tmp/subminer-socket\n',
    );

    const result = installFirstRunPluginToDefaultLocation({
      platform: 'win32',
      homeDir,
      dirname: path.join(root, 'dist', 'main', 'runtime'),
      appPath: path.join(root, 'app'),
      resourcesPath,
      binaryPath: 'C:\\Program Files\\SubMiner\\SubMiner.exe',
    });

    assert.equal(result.ok, true);
    assert.equal(
      fs.readFileSync(installPaths.pluginConfigPath, 'utf8'),
      'binary_path=C:\\Program Files\\SubMiner\\SubMiner.exe\nsocket_path=\\\\.\\pipe\\subminer-socket\n',
    );
  });
});

test('syncInstalledFirstRunPluginBinaryPath fills blank binary_path for existing installs', () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const xdgConfigHome = path.join(root, 'xdg');
    const installPaths = resolveDefaultMpvInstallPaths('linux', homeDir, xdgConfigHome);

    fs.mkdirSync(path.dirname(installPaths.pluginConfigPath), { recursive: true });
    fs.writeFileSync(
      installPaths.pluginConfigPath,
      'binary_path=\nsocket_path=/tmp/subminer-socket\n',
    );

    const result = syncInstalledFirstRunPluginBinaryPath({
      platform: 'linux',
      homeDir,
      xdgConfigHome,
      binaryPath: '/Applications/SubMiner.app/Contents/MacOS/SubMiner',
    });

    assert.deepEqual(result, {
      updated: true,
      configPath: installPaths.pluginConfigPath,
    });
    assert.equal(
      fs.readFileSync(installPaths.pluginConfigPath, 'utf8'),
      'binary_path=/Applications/SubMiner.app/Contents/MacOS/SubMiner\nsocket_path=/tmp/subminer-socket\n',
    );
  });
});

test('syncInstalledFirstRunPluginBinaryPath preserves explicit binary_path overrides', () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const xdgConfigHome = path.join(root, 'xdg');
    const installPaths = resolveDefaultMpvInstallPaths('linux', homeDir, xdgConfigHome);

    fs.mkdirSync(path.dirname(installPaths.pluginConfigPath), { recursive: true });
    fs.writeFileSync(
      installPaths.pluginConfigPath,
      'binary_path=/tmp/SubMiner/scripts/subminer-dev.sh\nsocket_path=/tmp/subminer-socket\n',
    );

    const result = syncInstalledFirstRunPluginBinaryPath({
      platform: 'linux',
      homeDir,
      xdgConfigHome,
      binaryPath: '/Applications/SubMiner.app/Contents/MacOS/SubMiner',
    });

    assert.deepEqual(result, {
      updated: false,
      configPath: installPaths.pluginConfigPath,
    });
    assert.equal(
      fs.readFileSync(installPaths.pluginConfigPath, 'utf8'),
      'binary_path=/tmp/SubMiner/scripts/subminer-dev.sh\nsocket_path=/tmp/subminer-socket\n',
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

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  detectInstalledFirstRunPlugin,
  installFirstRunPluginToDefaultLocation,
  resolvePackagedFirstRunPluginAssets,
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
    });

    assert.equal(result.ok, true);
    assert.equal(result.pluginInstallStatus, 'installed');
    assert.equal(detectInstalledFirstRunPlugin(installPaths), true);
    assert.equal(
      fs.readFileSync(installPaths.pluginEntrypointPath, 'utf8'),
      '-- packaged plugin',
    );
    assert.equal(fs.readFileSync(installPaths.pluginConfigPath, 'utf8'), 'configured=true\n');

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
    });

    assert.equal(result.ok, true);
    assert.equal(result.pluginInstallStatus, 'installed');
    assert.equal(detectInstalledFirstRunPlugin(installPaths), true);
    assert.equal(
      fs.readFileSync(installPaths.pluginEntrypointPath, 'utf8'),
      '-- packaged plugin',
    );
    assert.equal(
      fs.readFileSync(installPaths.pluginConfigPath, 'utf8'),
      'configured=true\n',
    );
  });
});

test('installFirstRunPluginToDefaultLocation rewrites Windows plugin socket_path', () => {
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
    });

    assert.equal(result.ok, true);
    assert.equal(
      fs.readFileSync(installPaths.pluginConfigPath, 'utf8'),
      'binary_path=\nsocket_path=\\\\.\\pipe\\subminer-socket\n',
    );
  });
});

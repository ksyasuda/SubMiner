import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  ensureLinuxRuntimePluginAssets,
  resolveManagedLinuxRuntimePluginPaths,
} from './linux-runtime-plugin-assets';

async function withTempDir<T>(fn: (dir: string) => Promise<T> | T): Promise<T> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-linux-plugin-assets-test-'));
  try {
    return await fn(dir);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

test('resolveManagedLinuxRuntimePluginPaths resolves XDG data target paths', () => {
  const resolved = resolveManagedLinuxRuntimePluginPaths({
    homeDir: '/home/tester',
    xdgDataHome: '/tmp/xdg-data',
  });

  assert.deepEqual(resolved, {
    rootDir: '/tmp/xdg-data/SubMiner/plugin',
    pluginDir: '/tmp/xdg-data/SubMiner/plugin/subminer',
    pluginEntrypointPath: '/tmp/xdg-data/SubMiner/plugin/subminer/main.lua',
    pluginConfigPath: '/tmp/xdg-data/SubMiner/plugin/subminer.conf',
  });
});

test('ensureLinuxRuntimePluginAssets installs managed plugin dir and config when missing', async () => {
  await withTempDir(async (tempDir) => {
    const sourceRoot = path.join(tempDir, 'source', 'plugin');
    const targetRoot = path.join(tempDir, 'xdg-data', 'SubMiner', 'plugin');
    fs.mkdirSync(path.join(sourceRoot, 'subminer'), { recursive: true });
    fs.writeFileSync(path.join(sourceRoot, 'subminer', 'main.lua'), '-- plugin\n');
    fs.writeFileSync(path.join(sourceRoot, 'subminer.conf'), 'configured=true\n');

    const result = await ensureLinuxRuntimePluginAssets({
      platform: 'linux',
      homeDir: path.join(tempDir, 'home'),
      xdgDataHome: path.join(tempDir, 'xdg-data'),
      resolveBundledAssets: () => ({
        pluginDirSource: path.join(sourceRoot, 'subminer'),
        pluginConfigSource: path.join(sourceRoot, 'subminer.conf'),
      }),
    });

    assert.deepEqual(result, {
      ok: true,
      status: 'installed',
      path: path.join(targetRoot, 'subminer', 'main.lua'),
    });
    assert.equal(
      fs.readFileSync(path.join(targetRoot, 'subminer', 'main.lua'), 'utf8'),
      '-- plugin\n',
    );
    assert.equal(
      fs.readFileSync(path.join(targetRoot, 'subminer.conf'), 'utf8'),
      'configured=true\n',
    );
  });
});

test('ensureLinuxRuntimePluginAssets returns already-present when managed assets already exist', async () => {
  await withTempDir(async (tempDir) => {
    const xdgDataHome = path.join(tempDir, 'xdg-data');
    const targetRoot = path.join(xdgDataHome, 'SubMiner', 'plugin');
    fs.mkdirSync(path.join(targetRoot, 'subminer'), { recursive: true });
    fs.writeFileSync(path.join(targetRoot, 'subminer', 'main.lua'), '-- existing\n');
    fs.writeFileSync(path.join(targetRoot, 'subminer.conf'), 'configured=true\n');

    const result = await ensureLinuxRuntimePluginAssets({
      platform: 'linux',
      homeDir: path.join(tempDir, 'home'),
      xdgDataHome,
      resolveBundledAssets: () => {
        throw new Error('should not resolve bundled assets when already installed');
      },
    });

    assert.deepEqual(result, {
      ok: true,
      status: 'already-present',
      path: path.join(targetRoot, 'subminer', 'main.lua'),
    });
  });
});

test('ensureLinuxRuntimePluginAssets fails when bundled assets cannot be resolved', async () => {
  await withTempDir(async (tempDir) => {
    const result = await ensureLinuxRuntimePluginAssets({
      platform: 'linux',
      homeDir: path.join(tempDir, 'home'),
      xdgDataHome: path.join(tempDir, 'xdg-data'),
      resolveBundledAssets: () => null,
    });

    assert.equal(result.ok, false);
    assert.equal(result.status, 'failed');
    assert.match(result.error ?? '', /bundled.*plugin assets/i);
  });
});

test('ensureLinuxRuntimePluginAssets leaves no final target tree on failed install', async () => {
  await withTempDir(async (tempDir) => {
    const sourceRoot = path.join(tempDir, 'source', 'plugin');
    const xdgDataHome = path.join(tempDir, 'xdg-data');
    const targetRoot = path.join(xdgDataHome, 'SubMiner', 'plugin');
    fs.mkdirSync(path.join(sourceRoot, 'subminer'), { recursive: true });
    fs.writeFileSync(path.join(sourceRoot, 'subminer', 'main.lua'), '-- plugin\n');
    fs.writeFileSync(path.join(sourceRoot, 'subminer.conf'), 'configured=true\n');

    const result = await ensureLinuxRuntimePluginAssets({
      platform: 'linux',
      homeDir: path.join(tempDir, 'home'),
      xdgDataHome,
      resolveBundledAssets: () => ({
        pluginDirSource: path.join(sourceRoot, 'subminer'),
        pluginConfigSource: path.join(sourceRoot, 'subminer.conf'),
      }),
      copyFile: async () => {
        throw new Error('copy failed');
      },
    });

    assert.equal(result.ok, false);
    assert.equal(result.status, 'failed');
    assert.equal(fs.existsSync(path.join(targetRoot, 'subminer', 'main.lua')), false);
    assert.equal(fs.existsSync(path.join(targetRoot, 'subminer.conf')), false);
  });
});

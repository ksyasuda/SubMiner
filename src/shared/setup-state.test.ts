import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  createDefaultSetupState,
  ensureDefaultConfigBootstrap,
  getDefaultConfigDir,
  getDefaultConfigFilePaths,
  getSetupStatePath,
  readSetupState,
  resolveDefaultMpvInstallPaths,
  writeSetupState,
} from './setup-state';

function withTempDir(fn: (dir: string) => void): void {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-setup-state-test-'));
  try {
    fn(dir);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

test('getDefaultConfigDir prefers existing SubMiner config directory', () => {
  const dir = getDefaultConfigDir({
    xdgConfigHome: '/tmp/xdg',
    homeDir: '/tmp/home',
    existsSync: (candidate) => candidate === '/tmp/xdg/SubMiner/config.jsonc',
  });

  assert.equal(dir, '/tmp/xdg/SubMiner');
});

test('ensureDefaultConfigBootstrap creates config dir and default jsonc only when missing', () => {
  withTempDir((root) => {
    const configDir = path.join(root, 'SubMiner');
    ensureDefaultConfigBootstrap({
      configDir,
      configFilePaths: getDefaultConfigFilePaths(configDir),
      generateTemplate: () => '{\n  "logging": {}\n}\n',
    });

    assert.equal(fs.existsSync(configDir), true);
    assert.equal(fs.readFileSync(path.join(configDir, 'config.jsonc'), 'utf8'), '{\n  "logging": {}\n}\n');

    fs.writeFileSync(path.join(configDir, 'config.json'), '{"keep":true}\n');
    fs.rmSync(path.join(configDir, 'config.jsonc'));
    ensureDefaultConfigBootstrap({
      configDir,
      configFilePaths: getDefaultConfigFilePaths(configDir),
      generateTemplate: () => 'should-not-write',
    });

    assert.equal(fs.existsSync(path.join(configDir, 'config.jsonc')), false);
    assert.equal(fs.readFileSync(path.join(configDir, 'config.json'), 'utf8'), '{"keep":true}\n');
  });
});

test('readSetupState ignores invalid files and round-trips valid state', () => {
  withTempDir((root) => {
    const statePath = getSetupStatePath(root);
    fs.writeFileSync(statePath, '{invalid');
    assert.equal(readSetupState(statePath), null);

    const state = createDefaultSetupState();
    state.status = 'completed';
    state.completionSource = 'user';
    state.lastSeenYomitanDictionaryCount = 2;
    writeSetupState(statePath, state);

    assert.deepEqual(readSetupState(statePath), state);
  });
});

test('resolveDefaultMpvInstallPaths resolves linux and macOS defaults', () => {
  assert.deepEqual(resolveDefaultMpvInstallPaths('linux', '/tmp/home', '/tmp/xdg'), {
    supported: true,
    mpvConfigDir: '/tmp/xdg/mpv',
    scriptsDir: '/tmp/xdg/mpv/scripts',
    scriptOptsDir: '/tmp/xdg/mpv/script-opts',
    pluginDir: '/tmp/xdg/mpv/scripts/subminer',
    pluginConfigPath: '/tmp/xdg/mpv/script-opts/subminer.conf',
  });

  assert.deepEqual(resolveDefaultMpvInstallPaths('darwin', '/Users/tester', undefined), {
    supported: true,
    mpvConfigDir: '/Users/tester/Library/Application Support/mpv',
    scriptsDir: '/Users/tester/Library/Application Support/mpv/scripts',
    scriptOptsDir: '/Users/tester/Library/Application Support/mpv/script-opts',
    pluginDir: '/Users/tester/Library/Application Support/mpv/scripts/subminer',
    pluginConfigPath: '/Users/tester/Library/Application Support/mpv/script-opts/subminer.conf',
  });
});

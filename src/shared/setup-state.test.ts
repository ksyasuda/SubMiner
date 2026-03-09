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
  const xdgConfigHome = path.join(path.sep, 'tmp', 'xdg');
  const homeDir = path.join(path.sep, 'tmp', 'home');
  const dir = getDefaultConfigDir({
    platform: 'linux',
    xdgConfigHome,
    homeDir,
    existsSync: (candidate) => candidate === path.join(xdgConfigHome, 'SubMiner', 'config.jsonc'),
  });

  assert.equal(dir, path.join(xdgConfigHome, 'SubMiner'));
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
    assert.equal(
      fs.readFileSync(path.join(configDir, 'config.jsonc'), 'utf8'),
      '{\n  "logging": {}\n}\n',
    );

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

test('ensureDefaultConfigBootstrap does not seed default config into an existing config directory', () => {
  withTempDir((root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'existing-user-file.txt'), 'keep\n');

    ensureDefaultConfigBootstrap({
      configDir,
      configFilePaths: getDefaultConfigFilePaths(configDir),
      generateTemplate: () => 'should-not-write',
    });

    assert.equal(fs.existsSync(path.join(configDir, 'config.jsonc')), false);
    assert.equal(
      fs.readFileSync(path.join(configDir, 'existing-user-file.txt'), 'utf8'),
      'keep\n',
    );
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

test('readSetupState migrates v1 state to v2 windows shortcut defaults', () => {
  withTempDir((root) => {
    const statePath = getSetupStatePath(root);
    fs.writeFileSync(
      statePath,
      JSON.stringify({
        version: 1,
        status: 'incomplete',
        completedAt: null,
        completionSource: null,
        lastSeenYomitanDictionaryCount: 0,
        pluginInstallStatus: 'unknown',
        pluginInstallPathSummary: null,
      }),
    );

    assert.deepEqual(readSetupState(statePath), {
      version: 2,
      status: 'incomplete',
      completedAt: null,
      completionSource: null,
      lastSeenYomitanDictionaryCount: 0,
      pluginInstallStatus: 'unknown',
      pluginInstallPathSummary: null,
      windowsMpvShortcutPreferences: {
        startMenuEnabled: true,
        desktopEnabled: true,
      },
      windowsMpvShortcutLastStatus: 'unknown',
    });
  });
});

test('resolveDefaultMpvInstallPaths resolves linux, macOS, and Windows defaults', () => {
  const linuxHomeDir = path.join(path.sep, 'tmp', 'home');
  const xdgConfigHome = path.join(path.sep, 'tmp', 'xdg');
  assert.deepEqual(resolveDefaultMpvInstallPaths('linux', linuxHomeDir, xdgConfigHome), {
    supported: true,
    mpvConfigDir: path.join(xdgConfigHome, 'mpv'),
    scriptsDir: path.join(xdgConfigHome, 'mpv', 'scripts'),
    scriptOptsDir: path.join(xdgConfigHome, 'mpv', 'script-opts'),
    pluginEntrypointPath: path.join(xdgConfigHome, 'mpv', 'scripts', 'subminer', 'main.lua'),
    pluginDir: path.join(xdgConfigHome, 'mpv', 'scripts', 'subminer'),
    pluginConfigPath: path.join(xdgConfigHome, 'mpv', 'script-opts', 'subminer.conf'),
  });

  const macHomeDir = path.join(path.sep, 'Users', 'tester');
  assert.deepEqual(resolveDefaultMpvInstallPaths('darwin', macHomeDir, undefined), {
    supported: true,
    mpvConfigDir: path.join(macHomeDir, 'Library', 'Application Support', 'mpv'),
    scriptsDir: path.join(macHomeDir, 'Library', 'Application Support', 'mpv', 'scripts'),
    scriptOptsDir: path.join(
      macHomeDir,
      'Library',
      'Application Support',
      'mpv',
      'script-opts',
    ),
    pluginEntrypointPath: path.join(
      macHomeDir,
      'Library',
      'Application Support',
      'mpv',
      'scripts',
      'subminer',
      'main.lua',
    ),
    pluginDir: path.join(
      macHomeDir,
      'Library',
      'Application Support',
      'mpv',
      'scripts',
      'subminer',
    ),
    pluginConfigPath: path.join(
      macHomeDir,
      'Library',
      'Application Support',
      'mpv',
      'script-opts',
      'subminer.conf',
    ),
  });

  assert.deepEqual(resolveDefaultMpvInstallPaths('win32', 'C:\\Users\\tester', undefined), {
    supported: true,
    mpvConfigDir: path.join('C:\\Users\\tester', 'AppData', 'Roaming', 'mpv'),
    scriptsDir: path.join('C:\\Users\\tester', 'AppData', 'Roaming', 'mpv', 'scripts'),
    scriptOptsDir: path.join('C:\\Users\\tester', 'AppData', 'Roaming', 'mpv', 'script-opts'),
    pluginEntrypointPath: path.join(
      'C:\\Users\\tester',
      'AppData',
      'Roaming',
      'mpv',
      'scripts',
      'subminer',
      'main.lua',
    ),
    pluginDir: path.join('C:\\Users\\tester', 'AppData', 'Roaming', 'mpv', 'scripts', 'subminer'),
    pluginConfigPath: path.join(
      'C:\\Users\\tester',
      'AppData',
      'Roaming',
      'mpv',
      'script-opts',
      'subminer.conf',
    ),
  });
});

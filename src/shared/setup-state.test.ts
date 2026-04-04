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
    existsSync: (candidate) =>
      candidate === path.posix.join(xdgConfigHome, 'SubMiner', 'config.jsonc'),
  });

  assert.equal(dir, path.posix.join(xdgConfigHome, 'SubMiner'));
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

test('ensureDefaultConfigBootstrap seeds default config into an existing config directory when missing', () => {
  withTempDir((root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'existing-user-file.txt'), 'keep\n');

    ensureDefaultConfigBootstrap({
      configDir,
      configFilePaths: getDefaultConfigFilePaths(configDir),
      generateTemplate: () => '{\n  "logging": {}\n}\n',
    });

    assert.equal(
      fs.readFileSync(path.join(configDir, 'config.jsonc'), 'utf8'),
      '{\n  "logging": {}\n}\n',
    );
    assert.equal(fs.readFileSync(path.join(configDir, 'existing-user-file.txt'), 'utf8'), 'keep\n');
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
    state.yomitanSetupMode = 'internal';
    state.lastSeenYomitanDictionaryCount = 2;
    writeSetupState(statePath, state);

    assert.deepEqual(readSetupState(statePath), state);
  });
});

test('readSetupState migrates v1 state to v3 windows shortcut defaults', () => {
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
      version: 3,
      status: 'incomplete',
      completedAt: null,
      completionSource: null,
      yomitanSetupMode: null,
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

test('readSetupState migrates completed v2 state to internal yomitan setup mode', () => {
  withTempDir((root) => {
    const statePath = getSetupStatePath(root);
    fs.writeFileSync(
      statePath,
      JSON.stringify({
        version: 2,
        status: 'completed',
        completedAt: '2026-03-12T00:00:00.000Z',
        completionSource: 'user',
        lastSeenYomitanDictionaryCount: 1,
        pluginInstallStatus: 'unknown',
        pluginInstallPathSummary: null,
        windowsMpvShortcutPreferences: {
          startMenuEnabled: true,
          desktopEnabled: true,
        },
        windowsMpvShortcutLastStatus: 'unknown',
      }),
    );

    assert.deepEqual(readSetupState(statePath), {
      version: 3,
      status: 'completed',
      completedAt: '2026-03-12T00:00:00.000Z',
      completionSource: 'user',
      yomitanSetupMode: 'internal',
      lastSeenYomitanDictionaryCount: 1,
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
    mpvConfigDir: path.posix.join(xdgConfigHome, 'mpv'),
    scriptsDir: path.posix.join(xdgConfigHome, 'mpv', 'scripts'),
    scriptOptsDir: path.posix.join(xdgConfigHome, 'mpv', 'script-opts'),
    pluginEntrypointPath: path.posix.join(xdgConfigHome, 'mpv', 'scripts', 'subminer', 'main.lua'),
    pluginDir: path.posix.join(xdgConfigHome, 'mpv', 'scripts', 'subminer'),
    pluginConfigPath: path.posix.join(xdgConfigHome, 'mpv', 'script-opts', 'subminer.conf'),
  });

  const macHomeDir = path.join(path.sep, 'Users', 'tester');
  assert.deepEqual(resolveDefaultMpvInstallPaths('darwin', macHomeDir, undefined), {
    supported: true,
    mpvConfigDir: path.posix.join(macHomeDir, '.config', 'mpv'),
    scriptsDir: path.posix.join(macHomeDir, '.config', 'mpv', 'scripts'),
    scriptOptsDir: path.posix.join(macHomeDir, '.config', 'mpv', 'script-opts'),
    pluginEntrypointPath: path.posix.join(
      macHomeDir,
      '.config',
      'mpv',
      'scripts',
      'subminer',
      'main.lua',
    ),
    pluginDir: path.posix.join(macHomeDir, '.config', 'mpv', 'scripts', 'subminer'),
    pluginConfigPath: path.posix.join(
      macHomeDir,
      '.config',
      'mpv',
      'script-opts',
      'subminer.conf',
    ),
  });

  assert.deepEqual(resolveDefaultMpvInstallPaths('win32', 'C:\\Users\\tester', undefined), {
    supported: true,
    mpvConfigDir: path.win32.join('C:\\Users\\tester', 'AppData', 'Roaming', 'mpv'),
    scriptsDir: path.win32.join('C:\\Users\\tester', 'AppData', 'Roaming', 'mpv', 'scripts'),
    scriptOptsDir: path.win32.join('C:\\Users\\tester', 'AppData', 'Roaming', 'mpv', 'script-opts'),
    pluginEntrypointPath: path.win32.join(
      'C:\\Users\\tester',
      'AppData',
      'Roaming',
      'mpv',
      'scripts',
      'subminer',
      'main.lua',
    ),
    pluginDir: path.win32.join(
      'C:\\Users\\tester',
      'AppData',
      'Roaming',
      'mpv',
      'scripts',
      'subminer',
    ),
    pluginConfigPath: path.win32.join(
      'C:\\Users\\tester',
      'AppData',
      'Roaming',
      'mpv',
      'script-opts',
      'subminer.conf',
    ),
  });
});

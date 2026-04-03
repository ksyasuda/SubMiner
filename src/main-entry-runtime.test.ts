import assert from 'node:assert/strict';
import test from 'node:test';
import {
  configureEarlyAppPaths,
  normalizeLaunchMpvExtraArgs,
  normalizeStartupArgv,
  normalizeLaunchMpvTargets,
  sanitizeHelpEnv,
  sanitizeLaunchMpvEnv,
  sanitizeStartupEnv,
  sanitizeBackgroundEnv,
  shouldDetachBackgroundLaunch,
  shouldHandleHelpOnlyAtEntry,
  shouldHandleLaunchMpvAtEntry,
  shouldHandleStatsDaemonCommandAtEntry,
} from './main-entry-runtime';

test('normalizeStartupArgv defaults no-arg startup to --start --background on non-Windows', () => {
  const originalPlatform = process.platform;
  try {
    Object.defineProperty(process, 'platform', { value: 'linux', configurable: true });

    assert.deepEqual(normalizeStartupArgv(['SubMiner.AppImage'], {}), [
      'SubMiner.AppImage',
      '--start',
      '--background',
    ]);
    assert.deepEqual(
      normalizeStartupArgv(['SubMiner.AppImage', '--password-store', 'gnome-libsecret'], {}),
      ['SubMiner.AppImage', '--password-store', 'gnome-libsecret', '--start', '--background'],
    );
    assert.deepEqual(normalizeStartupArgv(['SubMiner.AppImage', '--background'], {}), [
      'SubMiner.AppImage',
      '--background',
      '--start',
    ]);
    assert.deepEqual(normalizeStartupArgv(['SubMiner.AppImage', '--help'], {}), [
      'SubMiner.AppImage',
      '--help',
    ]);
  } finally {
    Object.defineProperty(process, 'platform', { value: originalPlatform, configurable: true });
  }
});

test('normalizeStartupArgv defaults no-arg Windows startup to --start only', () => {
  const originalPlatform = process.platform;
  try {
    Object.defineProperty(process, 'platform', { value: 'win32', configurable: true });

    assert.deepEqual(normalizeStartupArgv(['SubMiner.exe'], {}), ['SubMiner.exe', '--start']);
  } finally {
    Object.defineProperty(process, 'platform', { value: originalPlatform, configurable: true });
  }
});

test('shouldHandleHelpOnlyAtEntry detects help-only invocation', () => {
  assert.equal(shouldHandleHelpOnlyAtEntry(['--help'], {}), true);
  assert.equal(shouldHandleHelpOnlyAtEntry(['--help', '--start'], {}), false);
  assert.equal(shouldHandleHelpOnlyAtEntry(['--start'], {}), false);
  assert.equal(shouldHandleHelpOnlyAtEntry(['--help'], { ELECTRON_RUN_AS_NODE: '1' }), false);
});

test('launch-mpv entry helpers detect and normalize targets', () => {
  assert.equal(shouldHandleLaunchMpvAtEntry(['SubMiner.exe', '--launch-mpv'], {}), true);
  assert.equal(
    shouldHandleLaunchMpvAtEntry(['SubMiner.exe', '--launch-mpv'], { ELECTRON_RUN_AS_NODE: '1' }),
    false,
  );
  assert.deepEqual(normalizeLaunchMpvTargets(['SubMiner.exe', '--launch-mpv']), []);
  assert.deepEqual(normalizeLaunchMpvTargets(['SubMiner.exe', '--launch-mpv', 'C:\\a.mkv']), [
    'C:\\a.mkv',
  ]);
  assert.deepEqual(
    normalizeLaunchMpvExtraArgs([
      'SubMiner.exe',
      '--launch-mpv',
      '--sub-file',
      'track.srt',
      'C:\\a.mkv',
    ]),
    ['--sub-file', 'track.srt'],
  );
  assert.deepEqual(
    normalizeLaunchMpvTargets([
      'SubMiner.exe',
      '--launch-mpv',
      '--sub-file',
      'track.srt',
      'C:\\a.mkv',
    ]),
    ['C:\\a.mkv'],
  );
  assert.deepEqual(
    normalizeLaunchMpvExtraArgs([
      'SubMiner.exe',
      '--launch-mpv',
      '--profile=subminer',
      '--pause=yes',
      'C:\\a.mkv',
    ]),
    ['--profile=subminer', '--pause=yes'],
  );
  assert.deepEqual(
    normalizeLaunchMpvExtraArgs([
      'SubMiner.exe',
      '--launch-mpv',
      '--input-ipc-server',
      '\\\\.\\pipe\\custom-subminer-socket',
      '--alang',
      'ja,jpn',
      'C:\\a.mkv',
    ]),
    ['--input-ipc-server', '\\\\.\\pipe\\custom-subminer-socket', '--alang', 'ja,jpn'],
  );
  assert.deepEqual(
    normalizeLaunchMpvTargets([
      'SubMiner.exe',
      '--launch-mpv',
      '--input-ipc-server',
      '\\\\.\\pipe\\custom-subminer-socket',
      '--alang',
      'ja,jpn',
      'C:\\a.mkv',
      'C:\\b.mkv',
    ]),
    ['C:\\a.mkv', 'C:\\b.mkv'],
  );
});

test('stats-daemon entry helper detects internal daemon commands', () => {
  assert.equal(
    shouldHandleStatsDaemonCommandAtEntry(['SubMiner.AppImage', '--stats-daemon-start'], {}),
    true,
  );
  assert.equal(
    shouldHandleStatsDaemonCommandAtEntry(['SubMiner.AppImage', '--stats-daemon-stop'], {}),
    true,
  );
  assert.equal(
    shouldHandleStatsDaemonCommandAtEntry(['SubMiner.AppImage', '--stats-daemon-start'], {
      ELECTRON_RUN_AS_NODE: '1',
    }),
    false,
  );
  assert.equal(shouldHandleStatsDaemonCommandAtEntry(['SubMiner.AppImage', '--start'], {}), false);
});

test('sanitizeStartupEnv suppresses warnings and lsfg layer', () => {
  const env = sanitizeStartupEnv({
    VK_INSTANCE_LAYERS: 'foo:lsfg-vk:bar',
  });
  assert.equal(env.NODE_NO_WARNINGS, '1');
  assert.equal('VK_INSTANCE_LAYERS' in env, false);
});

test('sanitizeHelpEnv suppresses warnings and lsfg layer', () => {
  const env = sanitizeHelpEnv({
    VK_INSTANCE_LAYERS: 'foo:lsfg-vk:bar',
  });
  assert.equal(env.NODE_NO_WARNINGS, '1');
  assert.equal('VK_INSTANCE_LAYERS' in env, false);
});

test('sanitizeLaunchMpvEnv suppresses warnings and lsfg layer', () => {
  const env = sanitizeLaunchMpvEnv({
    VK_INSTANCE_LAYERS: 'foo:lsfg-vk:bar',
  });
  assert.equal(env.NODE_NO_WARNINGS, '1');
  assert.equal('VK_INSTANCE_LAYERS' in env, false);
});

test('sanitizeBackgroundEnv marks background child and keeps warning suppression', () => {
  const env = sanitizeBackgroundEnv({
    VK_INSTANCE_LAYERS: 'foo:lsfg-vk:bar',
  });
  assert.equal(env.SUBMINER_BACKGROUND_CHILD, '1');
  assert.equal(env.NODE_NO_WARNINGS, '1');
  assert.equal('VK_INSTANCE_LAYERS' in env, false);
});

test('shouldDetachBackgroundLaunch only for first background invocation', () => {
  assert.equal(shouldDetachBackgroundLaunch(['--background'], {}), true);
  assert.equal(
    shouldDetachBackgroundLaunch(['--background'], { SUBMINER_BACKGROUND_CHILD: '1' }),
    false,
  );
  assert.equal(
    shouldDetachBackgroundLaunch(['--background'], { ELECTRON_RUN_AS_NODE: '1' }),
    false,
  );
  assert.equal(shouldDetachBackgroundLaunch(['--start'], {}), false);
});

test('configureEarlyAppPaths pins userData to canonical SubMiner config dir', () => {
  const calls: string[] = [];

  const userDataPath = configureEarlyAppPaths(
    {
      setName: (name) => {
        calls.push(`name:${name}`);
      },
      setPath: (key, value) => {
        calls.push(`path:${key}:${value}`);
      },
    },
    {
      platform: 'linux',
      homeDir: '/home/tester',
      xdgConfigHome: '/tmp/xdg',
      existsSync: (candidate) => candidate === '/tmp/xdg/subminer/config.jsonc',
    },
  );

  assert.equal(userDataPath, '/tmp/xdg/SubMiner');
  assert.deepEqual(calls, ['name:SubMiner', 'path:userData:/tmp/xdg/SubMiner']);
});

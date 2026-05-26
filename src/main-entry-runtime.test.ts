import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { DEFAULT_CONFIG } from './config/definitions';
import { readConfiguredWindowsMpvLaunch } from './main-entry-launch-config';
import {
  configureEarlyAppPaths,
  normalizeLaunchMpvExtraArgs,
  normalizeStartupArgv,
  normalizeLaunchMpvTargets,
  resolveStatsDaemonCommandAction,
  sanitizeHelpEnv,
  sanitizeLaunchMpvEnv,
  sanitizeStartupEnv,
  sanitizeBackgroundEnv,
  shouldDetachBackgroundLaunch,
  shouldHandleHelpOnlyAtEntry,
  shouldHandleLaunchMpvAtEntry,
  shouldHandleStatsDaemonCommandAtEntry,
  hasTransportedStartupArgs,
  shouldForwardStartupArgvViaAppControl,
  applyEarlyLinuxCommandLineSwitches,
  resolveLinuxPasswordStoreValue,
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

test('normalizeStartupArgv uses transported AppImage args instead of raw Electron args', () => {
  assert.deepEqual(
    normalizeStartupArgv(['SubMiner.AppImage', '--background'], {
      SUBMINER_APP_ARGC: '2',
      SUBMINER_APP_ARG_0: '--stop',
      SUBMINER_APP_ARG_1: '--socket',
    }),
    ['SubMiner.AppImage', '--stop', '--socket'],
  );
});

test('normalizeStartupArgv defaults empty transported AppImage args to background startup', () => {
  const originalPlatform = process.platform;
  try {
    Object.defineProperty(process, 'platform', { value: 'linux', configurable: true });

    assert.deepEqual(
      normalizeStartupArgv(['SubMiner.AppImage', '--background'], {
        SUBMINER_APP_ARGC: '0',
      }),
      ['SubMiner.AppImage', '--start', '--background'],
    );
  } finally {
    Object.defineProperty(process, 'platform', { value: originalPlatform, configurable: true });
  }
});

test('normalizeStartupArgv defaults passive-only transported AppImage args to background startup', () => {
  const originalPlatform = process.platform;
  try {
    Object.defineProperty(process, 'platform', { value: 'linux', configurable: true });

    assert.deepEqual(
      normalizeStartupArgv(['SubMiner.AppImage'], {
        SUBMINER_APP_ARGC: '2',
        SUBMINER_APP_ARG_0: '--password-store',
        SUBMINER_APP_ARG_1: 'gnome-libsecret',
      }),
      ['SubMiner.AppImage', '--password-store', 'gnome-libsecret', '--start', '--background'],
    );
  } finally {
    Object.defineProperty(process, 'platform', { value: originalPlatform, configurable: true });
  }
});

test('hasTransportedStartupArgs detects env-carried app args', () => {
  assert.equal(hasTransportedStartupArgs({ SUBMINER_APP_ARGC: '1' }), true);
  assert.equal(hasTransportedStartupArgs({}), false);
});

test('resolveLinuxPasswordStoreValue defaults Linux safeStorage to gnome-libsecret', () => {
  assert.equal(resolveLinuxPasswordStoreValue(['SubMiner.AppImage'], 'linux'), 'gnome-libsecret');
  assert.equal(
    resolveLinuxPasswordStoreValue(['SubMiner.AppImage', '--password-store', 'gnome'], 'linux'),
    'gnome-libsecret',
  );
  assert.equal(resolveLinuxPasswordStoreValue(['SubMiner.exe'], 'win32'), null);
});

test('resolveLinuxPasswordStoreValue keeps scanning after a bare password-store flag', () => {
  assert.equal(
    resolveLinuxPasswordStoreValue(
      ['SubMiner.AppImage', '--password-store', '--start', '--password-store=kwallet6'],
      'linux',
    ),
    'kwallet6',
  );
});

test('applyEarlyLinuxCommandLineSwitches appends password store before main startup', () => {
  const switches: Array<[string, string | undefined]> = [];
  applyEarlyLinuxCommandLineSwitches(
    {
      appendSwitch: (name, value) => {
        switches.push([name, value]);
      },
    },
    ['SubMiner.AppImage', '--password-store=kwallet6'],
    'linux',
  );

  assert.deepEqual(switches, [
    ['enable-features', 'GlobalShortcutsPortal'],
    ['password-store', 'kwallet6'],
  ]);
});

test('transported AppImage visibility commands forward through app control', () => {
  assert.equal(
    shouldForwardStartupArgvViaAppControl(['SubMiner.AppImage', '--hide-visible-overlay'], {
      SUBMINER_APP_ARGC: '1',
      SUBMINER_APP_ARG_0: '--hide-visible-overlay',
    }),
    true,
  );
});

test('direct runtime commands forward through app control', () => {
  assert.equal(
    shouldForwardStartupArgvViaAppControl(['SubMiner.AppImage', '--hide-visible-overlay'], {}),
    true,
  );
  assert.equal(
    shouldForwardStartupArgvViaAppControl(
      ['SubMiner.exe', '--start', '--socket', '\\\\.\\pipe\\subminer-socket'],
      {},
    ),
    true,
  );
  assert.equal(shouldForwardStartupArgvViaAppControl(['SubMiner.exe', '--settings'], {}), true);
  assert.equal(shouldForwardStartupArgvViaAppControl(['SubMiner.exe', '--stop'], {}), true);
});

test('entry-only and internal commands do not forward through app control', () => {
  assert.equal(shouldForwardStartupArgvViaAppControl(['SubMiner.exe'], {}), false);
  assert.equal(shouldForwardStartupArgvViaAppControl(['SubMiner.exe', '--help'], {}), false);
  assert.equal(
    shouldForwardStartupArgvViaAppControl(['SubMiner.exe', '--generate-config'], {}),
    false,
  );
  assert.equal(
    shouldForwardStartupArgvViaAppControl(['SubMiner.exe', '--stats-daemon-start'], {}),
    false,
  );
  assert.equal(
    shouldForwardStartupArgvViaAppControl(['SubMiner.exe', '--stats', '--stats-background'], {}),
    false,
  );
  assert.equal(
    shouldForwardStartupArgvViaAppControl(['SubMiner.AppImage', '--app-ping'], {
      SUBMINER_APP_ARGC: '1',
      SUBMINER_APP_ARG_0: '--app-ping',
    }),
    false,
  );
  assert.equal(
    shouldForwardStartupArgvViaAppControl(['SubMiner.AppImage', '--launch-mpv'], {
      SUBMINER_APP_ARGC: '1',
      SUBMINER_APP_ARG_0: '--launch-mpv',
    }),
    false,
  );
  assert.equal(
    shouldForwardStartupArgvViaAppControl(['SubMiner.exe', '--start'], {
      ELECTRON_RUN_AS_NODE: '1',
    }),
    false,
  );
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
    normalizeLaunchMpvExtraArgs(['SubMiner.exe', '--launch-mpv', '--fullscreen', 'C:\\a.mkv']),
    ['--fullscreen'],
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
  assert.deepEqual(
    normalizeLaunchMpvTargets(['SubMiner.exe', '--launch-mpv', '--fullscreen', 'C:\\a.mkv']),
    ['C:\\a.mkv'],
  );
  assert.deepEqual(
    normalizeLaunchMpvExtraArgs([
      'SubMiner.exe',
      '--launch-mpv',
      '--msg-level',
      'all=warn',
      'C:\\a.mkv',
    ]),
    ['--msg-level', 'all=warn'],
  );
});

test('readConfiguredWindowsMpvLaunch includes defaults for runtime plugin script opts', () => {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-entry-config-'));
  try {
    const launch = readConfiguredWindowsMpvLaunch(tempDir);

    assert.equal(launch.executablePath, DEFAULT_CONFIG.mpv.executablePath);
    assert.equal(launch.launchMode, DEFAULT_CONFIG.mpv.launchMode);
    assert.equal(launch.logLevel, DEFAULT_CONFIG.logging.level);
    assert.equal(launch.logRotation, DEFAULT_CONFIG.logging.rotation);
    assert.deepEqual(launch.logFiles, DEFAULT_CONFIG.logging.files);
    assert.deepEqual(launch.pluginRuntimeConfig, {
      socketPath: DEFAULT_CONFIG.mpv.socketPath,
      binaryPath: DEFAULT_CONFIG.mpv.subminerBinaryPath,
      backend: DEFAULT_CONFIG.mpv.backend,
      logLevel: DEFAULT_CONFIG.logging.level,
      logRotation: DEFAULT_CONFIG.logging.rotation,
      autoStart: DEFAULT_CONFIG.mpv.autoStartSubMiner,
      autoStartVisibleOverlay: DEFAULT_CONFIG.auto_start_overlay,
      autoStartPauseUntilReady: DEFAULT_CONFIG.mpv.pauseUntilOverlayReady,
      texthookerEnabled: DEFAULT_CONFIG.texthooker.launchAtStartup,
      aniskipEnabled: DEFAULT_CONFIG.mpv.aniskipEnabled,
      aniskipButtonKey: DEFAULT_CONFIG.mpv.aniskipButtonKey,
    });
  } finally {
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
});

test('readConfiguredWindowsMpvLaunch preserves configured runtime plugin script opts', () => {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-entry-config-'));
  try {
    fs.writeFileSync(
      path.join(tempDir, 'config.jsonc'),
      JSON.stringify({
        auto_start_overlay: false,
        texthooker: {
          launchAtStartup: true,
        },
        logging: {
          level: 'debug',
          rotation: 14,
          files: {
            mpv: true,
          },
        },
        mpv: {
          executablePath: '  C:\\tools\\mpv.exe  ',
          launchMode: 'maximized',
          socketPath: '\\\\.\\pipe\\custom-subminer-socket',
          backend: 'windows',
          autoStartSubMiner: false,
          pauseUntilOverlayReady: false,
          subminerBinaryPath: 'C:\\SubMiner\\Custom.exe',
          aniskipEnabled: false,
          aniskipButtonKey: 'F8',
        },
      }),
    );

    const launch = readConfiguredWindowsMpvLaunch(tempDir);

    assert.equal(launch.executablePath, 'C:\\tools\\mpv.exe');
    assert.equal(launch.launchMode, 'maximized');
    assert.equal(launch.logLevel, 'debug');
    assert.equal(launch.logRotation, 14);
    assert.equal(launch.logFiles.mpv, true);
    assert.deepEqual(launch.pluginRuntimeConfig, {
      socketPath: '\\\\.\\pipe\\custom-subminer-socket',
      binaryPath: 'C:\\SubMiner\\Custom.exe',
      backend: 'windows',
      logLevel: 'debug',
      logRotation: 14,
      autoStart: false,
      autoStartVisibleOverlay: false,
      autoStartPauseUntilReady: false,
      texthookerEnabled: true,
      aniskipEnabled: false,
      aniskipButtonKey: 'F8',
    });
  } finally {
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
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

test('stats-daemon entry helper detects public background stats commands', () => {
  assert.equal(
    shouldHandleStatsDaemonCommandAtEntry(
      ['SubMiner.AppImage', '--stats', '--stats-background'],
      {},
    ),
    true,
  );
  assert.equal(
    shouldHandleStatsDaemonCommandAtEntry(['SubMiner.AppImage', '--stats', '--stats-stop'], {}),
    true,
  );
  assert.equal(
    shouldHandleStatsDaemonCommandAtEntry(['SubMiner.AppImage', '--stats-background'], {}),
    true,
  );
  assert.equal(
    shouldHandleStatsDaemonCommandAtEntry(['SubMiner.AppImage', '--stats-background'], {
      ELECTRON_RUN_AS_NODE: '1',
    }),
    false,
  );
  assert.equal(shouldHandleStatsDaemonCommandAtEntry(['SubMiner.AppImage', '--stats'], {}), false);
});

test('stats-daemon entry helper resolves daemon action for public and internal commands', () => {
  assert.equal(
    resolveStatsDaemonCommandAction(['SubMiner.AppImage', '--stats-daemon-start']),
    'start',
  );
  assert.equal(
    resolveStatsDaemonCommandAction(['SubMiner.AppImage', '--stats-daemon-stop']),
    'stop',
  );
  assert.equal(
    resolveStatsDaemonCommandAction(['SubMiner.AppImage', '--stats', '--stats-background']),
    'start',
  );
  assert.equal(
    resolveStatsDaemonCommandAction(['SubMiner.AppImage', '--stats', '--stats-stop']),
    'stop',
  );
  assert.equal(resolveStatsDaemonCommandAction(['SubMiner.AppImage', '--stats']), null);
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

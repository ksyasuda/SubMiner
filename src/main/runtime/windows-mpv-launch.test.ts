import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildWindowsMpvLaunchArgs,
  launchWindowsMpv,
  resolveWindowsMpvPath,
  type WindowsMpvLaunchDeps,
} from './windows-mpv-launch';

function createDeps(overrides: Partial<WindowsMpvLaunchDeps> = {}): WindowsMpvLaunchDeps {
  return {
    getEnv: () => undefined,
    runWhere: () => ({ status: 1, stdout: '' }),
    fileExists: () => false,
    spawnDetached: async () => undefined,
    showError: () => undefined,
    ...overrides,
  };
}

test('resolveWindowsMpvPath prefers SUBMINER_MPV_PATH', () => {
  const resolved = resolveWindowsMpvPath(
    createDeps({
      getEnv: (name) => (name === 'SUBMINER_MPV_PATH' ? 'C:\\mpv\\mpv.exe' : undefined),
      fileExists: (candidate) => candidate === 'C:\\mpv\\mpv.exe',
    }),
  );

  assert.equal(resolved, 'C:\\mpv\\mpv.exe');
});

test('resolveWindowsMpvPath prefers configured executable path before PATH', () => {
  const resolved = resolveWindowsMpvPath(
    createDeps({
      getEnv: () => undefined,
      runWhere: () => ({ status: 0, stdout: 'C:\\tools\\mpv.exe\r\n' }),
      fileExists: (candidate) => candidate === 'C:\\mpv\\mpv.exe',
    }),
    '  C:\\mpv\\mpv.exe  ',
  );

  assert.equal(resolved, 'C:\\mpv\\mpv.exe');
});

test('resolveWindowsMpvPath falls back to where.exe output', () => {
  const resolved = resolveWindowsMpvPath(
    createDeps({
      runWhere: () => ({ status: 0, stdout: 'C:\\tools\\mpv.exe\r\nC:\\other\\mpv.exe\r\n' }),
      fileExists: (candidate) => candidate === 'C:\\tools\\mpv.exe',
    }),
  );

  assert.equal(resolved, 'C:\\tools\\mpv.exe');
});

test('buildWindowsMpvLaunchArgs uses explicit SubMiner defaults and targets', () => {
  assert.deepEqual(
    buildWindowsMpvLaunchArgs(
      ['C:\\a.mkv', 'C:\\b.mkv'],
      [],
      'C:\\SubMiner\\SubMiner.exe',
      'C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main.lua',
    ),
    [
      '--player-operation-mode=pseudo-gui',
      '--force-window=immediate',
      '--script=C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main.lua',
      '--input-ipc-server=\\\\.\\pipe\\subminer-socket',
      '--alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
      '--slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
      '--sub-auto=fuzzy',
      '--sub-file-paths=subs;subtitles',
      '--sid=auto',
      '--secondary-sid=auto',
      '--sub-visibility=no',
      '--secondary-sub-visibility=no',
      '--script-opts=subminer-binary_path=C:\\SubMiner\\SubMiner.exe,subminer-socket_path=\\\\.\\pipe\\subminer-socket',
      'C:\\a.mkv',
      'C:\\b.mkv',
    ],
  );
});

test('buildWindowsMpvLaunchArgs inserts maximized launch mode before explicit extra args when configured', () => {
  assert.deepEqual(
    buildWindowsMpvLaunchArgs(
      ['C:\\video.mkv'],
      ['--window-maximized=no'],
      'C:\\SubMiner\\SubMiner.exe',
      'C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main.lua',
      'maximized',
    ),
    [
      '--player-operation-mode=pseudo-gui',
      '--force-window=immediate',
      '--script=C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main.lua',
      '--input-ipc-server=\\\\.\\pipe\\subminer-socket',
      '--alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
      '--slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
      '--sub-auto=fuzzy',
      '--sub-file-paths=subs;subtitles',
      '--sid=auto',
      '--secondary-sid=auto',
      '--sub-visibility=no',
      '--secondary-sub-visibility=no',
      '--script-opts=subminer-binary_path=C:\\SubMiner\\SubMiner.exe,subminer-socket_path=\\\\.\\pipe\\subminer-socket',
      '--window-maximized=yes',
      '--window-maximized=no',
      'C:\\video.mkv',
    ],
  );
});

test('buildWindowsMpvLaunchArgs keeps shortcut-only launches in idle mode', () => {
  assert.deepEqual(
    buildWindowsMpvLaunchArgs(
      [],
      [],
      'C:\\SubMiner\\SubMiner.exe',
      'C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main.lua',
    ),
    [
      '--player-operation-mode=pseudo-gui',
      '--force-window=immediate',
      '--idle=yes',
      '--script=C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main.lua',
      '--input-ipc-server=\\\\.\\pipe\\subminer-socket',
      '--alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
      '--slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
      '--sub-auto=fuzzy',
      '--sub-file-paths=subs;subtitles',
      '--sid=auto',
      '--secondary-sid=auto',
      '--sub-visibility=no',
      '--secondary-sub-visibility=no',
      '--script-opts=subminer-binary_path=C:\\SubMiner\\SubMiner.exe,subminer-socket_path=\\\\.\\pipe\\subminer-socket',
    ],
  );
});

test('buildWindowsMpvLaunchArgs mirrors a custom input-ipc-server into script opts', () => {
  assert.deepEqual(
    buildWindowsMpvLaunchArgs(
      ['C:\\video.mkv'],
      ['--input-ipc-server', '\\\\.\\pipe\\custom-subminer-socket'],
      'C:\\SubMiner\\SubMiner.exe',
      'C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main.lua',
    ),
    [
      '--player-operation-mode=pseudo-gui',
      '--force-window=immediate',
      '--script=C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main.lua',
      '--input-ipc-server=\\\\.\\pipe\\custom-subminer-socket',
      '--alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
      '--slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
      '--sub-auto=fuzzy',
      '--sub-file-paths=subs;subtitles',
      '--sid=auto',
      '--secondary-sid=auto',
      '--sub-visibility=no',
      '--secondary-sub-visibility=no',
      '--script-opts=subminer-binary_path=C:\\SubMiner\\SubMiner.exe,subminer-socket_path=\\\\.\\pipe\\custom-subminer-socket',
      '--input-ipc-server',
      '\\\\.\\pipe\\custom-subminer-socket',
      'C:\\video.mkv',
    ],
  );
});

test('buildWindowsMpvLaunchArgs includes socket script opts when plugin entrypoint is present without binary path', () => {
  assert.deepEqual(
    buildWindowsMpvLaunchArgs(
      ['C:\\video.mkv'],
      ['--input-ipc-server', '\\\\.\\pipe\\custom-subminer-socket'],
      undefined,
      'C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main.lua',
    ),
    [
      '--player-operation-mode=pseudo-gui',
      '--force-window=immediate',
      '--script=C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main.lua',
      '--input-ipc-server=\\\\.\\pipe\\custom-subminer-socket',
      '--alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
      '--slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
      '--sub-auto=fuzzy',
      '--sub-file-paths=subs;subtitles',
      '--sid=auto',
      '--secondary-sid=auto',
      '--sub-visibility=no',
      '--secondary-sub-visibility=no',
      '--script-opts=subminer-socket_path=\\\\.\\pipe\\custom-subminer-socket',
      '--input-ipc-server',
      '\\\\.\\pipe\\custom-subminer-socket',
      'C:\\video.mkv',
    ],
  );
});

test('buildWindowsMpvLaunchArgs uses runtime plugin config script opts', () => {
  const args = buildWindowsMpvLaunchArgs(
    ['C:\\video.mkv'],
    ['--input-ipc-server', '\\\\.\\pipe\\custom-subminer-socket'],
    'C:\\SubMiner\\SubMiner.exe',
    'C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main.lua',
    'normal',
    {
      socketPath: '\\\\.\\pipe\\ignored-config-socket',
      binaryPath: 'C:\\Custom\\SubMiner.exe',
      backend: 'windows',
      autoStart: true,
      autoStartVisibleOverlay: false,
      autoStartPauseUntilReady: false,
      texthookerEnabled: false,
      aniskipEnabled: true,
      aniskipButtonKey: 'F8',
    },
  );

  const scriptOpts = args.find((arg) => arg.startsWith('--script-opts='));
  assert.match(scriptOpts ?? '', /subminer-binary_path=C:\\Custom\\SubMiner\.exe/);
  assert.match(scriptOpts ?? '', /subminer-socket_path=\\\\\.\\pipe\\custom-subminer-socket/);
  assert.match(scriptOpts ?? '', /subminer-backend=windows/);
  assert.match(scriptOpts ?? '', /subminer-auto_start=yes/);
  assert.match(scriptOpts ?? '', /subminer-auto_start_visible_overlay=no/);
  assert.match(scriptOpts ?? '', /subminer-auto_start_pause_until_ready=no/);
  assert.match(scriptOpts ?? '', /subminer-texthooker_enabled=no/);
  assert.match(scriptOpts ?? '', /subminer-aniskip_enabled=yes/);
  assert.match(scriptOpts ?? '', /subminer-aniskip_button_key=F8/);
});

test('buildWindowsMpvLaunchArgs keeps Windows ipc default unless explicitly overridden', () => {
  const args = buildWindowsMpvLaunchArgs(
    ['C:\\video.mkv'],
    [],
    'C:\\SubMiner\\SubMiner.exe',
    'C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main.lua',
    'normal',
    {
      socketPath: 'C:\\Users\\tester\\AppData\\Local\\Temp\\subminer-smoke-sock\\subminer.sock',
      binaryPath: '',
      backend: 'windows',
      autoStart: true,
      autoStartVisibleOverlay: true,
      autoStartPauseUntilReady: true,
      texthookerEnabled: false,
      aniskipEnabled: true,
      aniskipButtonKey: 'F7',
    },
  );

  assert.ok(args.includes('--input-ipc-server=\\\\.\\pipe\\subminer-socket'));
  const scriptOpts = args.find((arg) => arg.startsWith('--script-opts='));
  assert.match(scriptOpts ?? '', /subminer-socket_path=\\\\\.\\pipe\\subminer-socket/);
});

test('launchWindowsMpv attaches a launched video to a running app and disables plugin auto-start', async () => {
  const spawnedArgs: string[][] = [];
  const controlArgv: string[][] = [];
  const waitedSockets: Array<{ socketPath: string; timeoutMs: number }> = [];
  const logs: string[] = [];
  const result = await launchWindowsMpv(
    ['C:\\video.mkv'],
    createDeps({
      getEnv: (name) => (name === 'SUBMINER_MPV_PATH' ? 'C:\\mpv\\mpv.exe' : undefined),
      fileExists: (candidate) => candidate === 'C:\\mpv\\mpv.exe',
      isAppControlServerAvailable: async () => true,
      waitForSocketReady: async (socketPath, timeoutMs) => {
        waitedSockets.push({ socketPath, timeoutMs });
        return true;
      },
      sendAppControlCommand: async (argv) => {
        controlArgv.push(argv);
        return { ok: true };
      },
      logInfo: (message) => logs.push(message),
      spawnDetached: async (_command, args) => {
        spawnedArgs.push(args);
      },
    }),
    ['--input-ipc-server', '\\\\.\\pipe\\warm-subminer'],
    'C:\\SubMiner\\SubMiner.exe',
    'C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main.lua',
    '',
    'normal',
    undefined,
    {
      socketPath: '\\\\.\\pipe\\ignored-config-socket',
      binaryPath: '',
      backend: 'windows',
      logLevel: 'debug',
      logRotation: 7,
      autoStart: true,
      autoStartVisibleOverlay: true,
      autoStartPauseUntilReady: true,
      texthookerEnabled: true,
      aniskipEnabled: true,
      aniskipButtonKey: 'TAB',
    },
  );

  assert.equal(result.ok, true);
  const scriptOpts = spawnedArgs[0]?.find((arg) => arg.startsWith('--script-opts='));
  assert.match(scriptOpts ?? '', /subminer-auto_start=no/);
  assert.match(scriptOpts ?? '', /subminer-socket_path=\\\\\.\\pipe\\warm-subminer/);
  assert.deepEqual(waitedSockets, [{ socketPath: '\\\\.\\pipe\\warm-subminer', timeoutMs: 10000 }]);
  assert.deepEqual(controlArgv, [
    [
      '--start',
      '--managed-playback',
      '--log-level',
      'debug',
      '--backend',
      'windows',
      '--socket',
      '\\\\.\\pipe\\warm-subminer',
      '--show-visible-overlay',
      '--texthooker',
    ],
  ]);
  assert.ok(logs.some((line) => line.includes('attachRunningApp=yes')));
  assert.ok(logs.some((line) => line.includes('Attached launched mpv session')));
});

test('launchWindowsMpv leaves plugin auto-start enabled when no running app control socket exists', async () => {
  const spawnedArgs: string[][] = [];
  let controlCalls = 0;
  let waitCalls = 0;
  const result = await launchWindowsMpv(
    ['C:\\video.mkv'],
    createDeps({
      getEnv: (name) => (name === 'SUBMINER_MPV_PATH' ? 'C:\\mpv\\mpv.exe' : undefined),
      fileExists: (candidate) => candidate === 'C:\\mpv\\mpv.exe',
      isAppControlServerAvailable: async () => false,
      waitForSocketReady: async () => {
        waitCalls += 1;
        return true;
      },
      sendAppControlCommand: async () => {
        controlCalls += 1;
        return { ok: true };
      },
      spawnDetached: async (_command, args) => {
        spawnedArgs.push(args);
      },
    }),
    [],
    'C:\\SubMiner\\SubMiner.exe',
    'C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main.lua',
    '',
    'normal',
    undefined,
    {
      socketPath: '\\\\.\\pipe\\subminer-socket',
      binaryPath: '',
      backend: 'windows',
      autoStart: true,
      autoStartVisibleOverlay: true,
      autoStartPauseUntilReady: true,
      texthookerEnabled: false,
      aniskipEnabled: true,
      aniskipButtonKey: 'TAB',
    },
  );

  assert.equal(result.ok, true);
  const scriptOpts = spawnedArgs[0]?.find((arg) => arg.startsWith('--script-opts='));
  assert.match(scriptOpts ?? '', /subminer-auto_start=yes/);
  assert.equal(waitCalls, 0);
  assert.equal(controlCalls, 0);
});

test('launchWindowsMpv reports missing mpv path', async () => {
  const errors: string[] = [];
  const result = await launchWindowsMpv(
    [],
    createDeps({
      showError: (_title, content) => errors.push(content),
    }),
  );

  assert.equal(result.ok, false);
  assert.equal(result.mpvPath, '');
  assert.match(errors[0] ?? '', /mpv\.executablePath/i);
});

test('launchWindowsMpv spawns detached mpv with targets', async () => {
  const calls: string[] = [];
  const logs: string[] = [];
  const result = await launchWindowsMpv(
    ['C:\\video.mkv'],
    createDeps({
      getEnv: (name) => (name === 'SUBMINER_MPV_PATH' ? 'C:\\mpv\\mpv.exe' : undefined),
      fileExists: (candidate) => candidate === 'C:\\mpv\\mpv.exe',
      logInfo: (message) => logs.push(message),
      spawnDetached: async (command, args) => {
        calls.push(command);
        calls.push(args.join('|'));
      },
    }),
    [],
    'C:\\SubMiner\\SubMiner.exe',
    'C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main.lua',
  );

  assert.equal(result.ok, true);
  assert.equal(result.mpvPath, 'C:\\mpv\\mpv.exe');
  assert.deepEqual(calls, [
    'C:\\mpv\\mpv.exe',
    '--player-operation-mode=pseudo-gui|--force-window=immediate|--script=C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main.lua|--input-ipc-server=\\\\.\\pipe\\subminer-socket|--alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us|--slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us|--sub-auto=fuzzy|--sub-file-paths=subs;subtitles|--sid=auto|--secondary-sid=auto|--sub-visibility=no|--secondary-sub-visibility=no|--script-opts=subminer-binary_path=C:\\SubMiner\\SubMiner.exe,subminer-socket_path=\\\\.\\pipe\\subminer-socket|C:\\video.mkv',
  ]);
  assert.match(logs[0] ?? '', /mpvPath=C:\\mpv\\mpv\.exe/);
  assert.match(logs[0] ?? '', /inputIpcServer=\\\\\.\\pipe\\subminer-socket/);
  assert.match(
    logs[0] ?? '',
    /bundledPlugin=C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main\.lua/,
  );
  assert.match(logs[0] ?? '', /installedPlugin=none/);
});

test('launchWindowsMpv forwards runtime logging config to mpv and plugin', async () => {
  const calls: string[] = [];
  const result = await launchWindowsMpv(
    ['C:\\video.mkv'],
    createDeps({
      getEnv: (name) => (name === 'SUBMINER_MPV_PATH' ? 'C:\\mpv\\mpv.exe' : undefined),
      fileExists: (candidate) => candidate === 'C:\\mpv\\mpv.exe',
      spawnDetached: async (command, args, env) => {
        calls.push(command);
        calls.push(args.join('|'));
        calls.push(env?.SUBMINER_LOG_LEVEL ?? '');
        calls.push(env?.SUBMINER_LOG_ROTATION ?? '');
      },
    }),
    ['--log-file=C:\\Users\\tester\\AppData\\Roaming\\SubMiner\\logs\\mpv-2026-05-26.log'],
    'C:\\SubMiner\\SubMiner.exe',
    'C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main.lua',
    '',
    'normal',
    undefined,
    {
      socketPath: '\\\\.\\pipe\\subminer-socket',
      binaryPath: '',
      backend: 'windows',
      logLevel: 'debug',
      logRotation: 0,
      autoStart: true,
      autoStartVisibleOverlay: false,
      autoStartPauseUntilReady: true,
      texthookerEnabled: false,
      aniskipEnabled: true,
      aniskipButtonKey: 'TAB',
    },
  );

  assert.equal(result.ok, true);
  assert.match(calls[1] ?? '', /--msg-level=all=warn,subminer=debug/);
  assert.doesNotMatch(calls[1] ?? '', /subminer-log_level=debug/);
  assert.match(
    calls[1] ?? '',
    /--log-file=C:\\Users\\tester\\AppData\\Roaming\\SubMiner\\logs\\mpv-2026-05-26\.log/,
  );
  assert.equal(calls[2], 'debug');
  assert.equal(calls[3], '0');
});

test('launchWindowsMpv skips bundled script when installed plugin is detected', async () => {
  const calls: string[] = [];
  const notifications: string[] = [];
  const result = await launchWindowsMpv(
    ['C:\\video.mkv'],
    createDeps({
      getEnv: (name) => (name === 'SUBMINER_MPV_PATH' ? 'C:\\mpv\\mpv.exe' : undefined),
      fileExists: (candidate) => candidate === 'C:\\mpv\\mpv.exe',
      spawnDetached: async (command, args) => {
        calls.push(command);
        calls.push(args.join('|'));
      },
    }),
    [],
    'C:\\SubMiner\\SubMiner.exe',
    'C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main.lua',
    '',
    'normal',
    {
      detectInstalledMpvPlugin: () => ({
        installed: true,
        path: 'C:\\Users\\tester\\AppData\\Roaming\\mpv\\scripts\\subminer\\main.lua',
        version: null,
        source: 'default-config',
        message: null,
      }),
      notifyInstalledPluginDetected: (detection) => {
        notifications.push(detection.path ?? '');
      },
    },
  );

  assert.equal(result.ok, true);
  assert.equal(calls[0], 'C:\\mpv\\mpv.exe');
  assert.doesNotMatch(calls[1] ?? '', /--script=C:\\Program Files\\SubMiner/);
  assert.match(calls[1] ?? '', /--script-opts=subminer-binary_path=C:\\SubMiner\\SubMiner\.exe/);
  assert.deepEqual(notifications, [
    'C:\\Users\\tester\\AppData\\Roaming\\mpv\\scripts\\subminer\\main.lua',
  ]);
});

test('launchWindowsMpv prompts before launch and injects bundled script after legacy plugin removal', async () => {
  const calls: string[] = [];
  const prompts: string[] = [];
  let detectCalls = 0;
  const result = await launchWindowsMpv(
    ['C:\\video.mkv'],
    createDeps({
      getEnv: (name) => (name === 'SUBMINER_MPV_PATH' ? 'C:\\mpv\\mpv.exe' : undefined),
      fileExists: (candidate) => candidate === 'C:\\mpv\\mpv.exe',
      spawnDetached: async (command, args) => {
        calls.push(command);
        calls.push(args.join('|'));
      },
    }),
    [],
    'C:\\SubMiner\\SubMiner.exe',
    'C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main.lua',
    '',
    'normal',
    {
      detectInstalledMpvPlugin: () => {
        detectCalls += 1;
        return detectCalls === 1
          ? {
              installed: true,
              path: 'C:\\Users\\tester\\AppData\\Roaming\\mpv\\scripts\\subminer\\main.lua',
              version: '0.12.0',
              source: 'default-config',
              message: null,
            }
          : {
              installed: false,
              path: null,
              version: null,
              source: null,
              message: null,
            };
      },
      resolveInstalledPluginBeforeLaunch: async (detection) => {
        prompts.push(detection.path ?? '');
        return 'removed' as const;
      },
    },
  );

  assert.equal(result.ok, true);
  assert.equal(detectCalls, 2);
  assert.deepEqual(prompts, [
    'C:\\Users\\tester\\AppData\\Roaming\\mpv\\scripts\\subminer\\main.lua',
  ]);
  assert.equal(calls[0], 'C:\\mpv\\mpv.exe');
  assert.match(
    calls[1] ?? '',
    /--script=C:\\Program Files\\SubMiner\\resources\\plugin\\subminer\\main\.lua/,
  );
});

test('launchWindowsMpv reports spawn failures with path context', async () => {
  const errors: string[] = [];
  const result = await launchWindowsMpv(
    [],
    createDeps({
      getEnv: (name) => (name === 'SUBMINER_MPV_PATH' ? 'C:\\mpv\\mpv.exe' : undefined),
      fileExists: (candidate) => candidate === 'C:\\mpv\\mpv.exe',
      spawnDetached: async () => {
        throw new Error('spawn failed');
      },
      showError: (_title, content) => errors.push(content),
    }),
  );

  assert.equal(result.ok, false);
  assert.equal(result.mpvPath, 'C:\\mpv\\mpv.exe');
  assert.match(errors[0] ?? '', /Failed to launch mpv/i);
  assert.match(errors[0] ?? '', /C:\\mpv\\mpv\.exe/i);
});

test('launchWindowsMpv reports async spawn failures with path context', async () => {
  const errors: string[] = [];
  const result = await launchWindowsMpv(
    [],
    createDeps({
      getEnv: (name) => (name === 'SUBMINER_MPV_PATH' ? 'C:\\mpv\\mpv.exe' : undefined),
      fileExists: (candidate) => candidate === 'C:\\mpv\\mpv.exe',
      spawnDetached: () => Promise.reject(new Error('async spawn failed')),
      showError: (_title, content) => errors.push(content),
    }),
  );

  assert.equal(result.ok, false);
  assert.equal(result.mpvPath, 'C:\\mpv\\mpv.exe');
  assert.match(errors[0] ?? '', /Failed to launch mpv/i);
  assert.match(errors[0] ?? '', /async spawn failed/i);
});

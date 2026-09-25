import test from 'node:test';
import assert from 'node:assert/strict';
import { detectInstalledMpvPlugin } from './first-run-setup-plugin';
import { resolveWindowsMpvPath } from './mpv-process';
import {
  createEnsureMpvConnectedForJellyfinPlaybackHandler,
  createLaunchMpvIdleForJellyfinPlaybackHandler,
  createWaitForMpvConnectedHandler,
} from './jellyfin-remote-connection';

test('createWaitForMpvConnectedHandler connects and waits for readiness', async () => {
  let connected = false;
  let nowMs = 0;
  const waitForConnected = createWaitForMpvConnectedHandler({
    getMpvClient: () => ({
      connected,
      connect: () => {
        connected = true;
      },
    }),
    now: () => nowMs,
    sleep: async () => {
      nowMs += 100;
    },
  });

  const ready = await waitForConnected(500);
  assert.equal(ready, true);
});

test('createLaunchMpvIdleForJellyfinPlaybackHandler builds expected mpv args', () => {
  const spawnedArgs: string[][] = [];
  const logs: string[] = [];
  const launch = createLaunchMpvIdleForJellyfinPlaybackHandler({
    getMpvExecutablePath: () => 'mpv',
    getSocketPath: () => '/tmp/subminer.sock',
    getLaunchMode: () => 'maximized',
    platform: 'darwin',
    execPath: '/Applications/SubMiner.app/Contents/MacOS/SubMiner',
    getRuntimePluginEntrypoint: () =>
      '/Applications/SubMiner.app/Contents/Resources/plugin/subminer/main.lua',
    getDefaultMpvLogPath: () => ' /tmp/mp.log ',
    defaultMpvArgs: ['--sid=auto'],
    removeSocketPath: () => {},
    spawnMpv: (_executable, args) => {
      spawnedArgs.push(args);
      return {
        on: () => {},
        unref: () => {},
      };
    },
    logWarn: (message) => logs.push(message),
    logInfo: (message) => logs.push(message),
  });

  launch();
  assert.equal(spawnedArgs.length, 1);
  assert.ok(spawnedArgs[0]!.includes('--window-maximized=yes'));
  assert.ok(spawnedArgs[0]!.includes('--idle=yes'));
  assert.ok(
    spawnedArgs[0]!.includes(
      '--script=/Applications/SubMiner.app/Contents/Resources/plugin/subminer/main.lua',
    ),
  );
  assert.ok(spawnedArgs[0]!.includes('--log-file=/tmp/mp.log'));
  assert.ok(spawnedArgs[0]!.some((arg) => arg.includes('--input-ipc-server=/tmp/subminer.sock')));
  assert.ok(logs.some((entry) => entry.includes('Launched mpv for Jellyfin playback')));
});

test('createLaunchMpvIdleForJellyfinPlaybackHandler forwards runtime plugin config', () => {
  const spawnedArgs: string[][] = [];
  const launch = createLaunchMpvIdleForJellyfinPlaybackHandler({
    getMpvExecutablePath: () => 'mpv',
    getSocketPath: () => '/tmp/subminer.sock',
    getLaunchMode: () => 'normal',
    platform: 'linux',
    execPath: '/opt/SubMiner/SubMiner.AppImage',
    getPluginRuntimeConfig: () => ({
      socketPath: '/tmp/ignored-config.sock',
      binaryPath: '/custom/SubMiner.AppImage',
      backend: 'x11',
      autoStart: true,
      autoStartVisibleOverlay: false,
      autoStartPauseUntilReady: false,
      osdMessages: false,
      texthookerEnabled: false,
    }),
    getDefaultMpvLogPath: () => '/tmp/mp.log',
    defaultMpvArgs: ['--sid=auto'],
    removeSocketPath: () => {},
    spawnMpv: (_executable, args) => {
      spawnedArgs.push(args);
      return {
        on: () => {},
        unref: () => {},
      };
    },
    logWarn: () => {},
    logInfo: () => {},
  });

  launch();
  const scriptOpts = spawnedArgs[0]?.find((arg) => arg.startsWith('--script-opts='));
  assert.match(scriptOpts ?? '', /subminer-binary_path=\/custom\/SubMiner\.AppImage/);
  assert.match(scriptOpts ?? '', /subminer-socket_path=\/tmp\/subminer\.sock/);
  assert.match(scriptOpts ?? '', /subminer-backend=x11/);
  assert.match(scriptOpts ?? '', /subminer-auto_start=yes/);
  assert.match(scriptOpts ?? '', /subminer-auto_start_visible_overlay=no/);
  assert.match(scriptOpts ?? '', /subminer-auto_start_pause_until_ready=no/);
  assert.match(scriptOpts ?? '', /subminer-texthooker_enabled=no/);
  assert.doesNotMatch(scriptOpts ?? '', /subminer-aniskip_enabled=/);
  assert.doesNotMatch(scriptOpts ?? '', /subminer-aniskip_button_key=/);
});

test('Jellyfin detects portable plugins beside the executable selected for launch', () => {
  const mpvPath = 'C:\\portable player\\mpv.exe';
  const pluginPath = 'C:\\portable player\\portable_config\\scripts\\subminer\\main.lua';
  for (const source of ['environment', 'PATH']) {
    let resolutions = 0;
    const spawned: Array<{ executable: string; args: string[] }> = [];
    const launch = createLaunchMpvIdleForJellyfinPlaybackHandler({
      getMpvExecutablePath: () => {
        resolutions += 1;
        return resolveWindowsMpvPath({
          getEnv: () => (source === 'environment' ? mpvPath : undefined),
          runWhere: () => ({ status: 0, stdout: mpvPath }),
          fileExists: (candidate) => candidate === mpvPath,
        });
      },
      getSocketPath: () => '\\\\.\\pipe\\subminer-test',
      getLaunchMode: () => 'normal',
      platform: 'win32',
      execPath: 'C:\\SubMiner\\SubMiner.exe',
      getRuntimePluginEntrypoint: () => 'C:\\SubMiner\\plugin\\subminer\\main.lua',
      getInstalledPluginDetection: (mpvExecutablePath) =>
        detectInstalledMpvPlugin({
          platform: 'win32',
          homeDir: 'C:\\Users\\test',
          mpvExecutablePath,
          existsSync: (candidate) => candidate === pluginPath,
        }),
      getDefaultMpvLogPath: () => '',
      defaultMpvArgs: [],
      removeSocketPath: () => {},
      spawnMpv: (executable, args) => {
        spawned.push({ executable, args });
        return { on: () => {}, unref: () => {} };
      },
      logWarn: () => {},
      logInfo: () => {},
    });

    launch();
    assert.equal(resolutions, 1, source);
    assert.equal(spawned.length, 1);
    assert.equal(spawned[0]!.executable, mpvPath);
    assert.equal(
      spawned[0]!.args.some((arg) => arg.startsWith('--script=')),
      false,
    );
  }
});

test('createEnsureMpvConnectedForJellyfinPlaybackHandler auto-launches once', async () => {
  let autoLaunchInFlight: Promise<boolean> | null = null;
  let launchCalls = 0;
  let waitCalls = 0;
  let mpvClient: { connected: boolean; connect: () => void } | null = null;
  let resolveAutoLaunchPromise: (value: boolean) => void = () => {};
  const autoLaunchPromise = new Promise<boolean>((resolve) => {
    resolveAutoLaunchPromise = resolve;
  });

  const ensureConnected = createEnsureMpvConnectedForJellyfinPlaybackHandler({
    getMpvClient: () => mpvClient,
    setMpvClient: (client) => {
      mpvClient = client;
    },
    createMpvClient: () => ({
      connected: false,
      connect: () => {},
    }),
    waitForMpvConnected: async (timeoutMs) => {
      waitCalls += 1;
      if (timeoutMs === 3000) return false;
      return await autoLaunchPromise;
    },
    launchMpvIdleForJellyfinPlayback: () => {
      launchCalls += 1;
    },
    getAutoLaunchInFlight: () => autoLaunchInFlight,
    setAutoLaunchInFlight: (promise) => {
      autoLaunchInFlight = promise;
    },
    connectTimeoutMs: 3000,
    autoLaunchTimeoutMs: 20000,
  });

  const firstPromise = ensureConnected();
  const secondPromise = ensureConnected();
  resolveAutoLaunchPromise(true);
  const first = await firstPromise;
  const second = await secondPromise;

  assert.equal(first, true);
  assert.equal(second, true);
  assert.equal(launchCalls, 1);
  assert.equal(waitCalls >= 2, true);
});

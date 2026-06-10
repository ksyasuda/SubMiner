import test from 'node:test';
import assert from 'node:assert/strict';
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
    getSocketPath: () => '/tmp/subminer.sock',
    getLaunchMode: () => 'maximized',
    platform: 'darwin',
    execPath: '/Applications/SubMiner.app/Contents/MacOS/SubMiner',
    getRuntimePluginEntrypoint: () =>
      '/Applications/SubMiner.app/Contents/Resources/plugin/subminer/main.lua',
    getDefaultMpvLogPath: () => ' /tmp/mp.log ',
    defaultMpvArgs: ['--sid=auto'],
    removeSocketPath: () => {},
    spawnMpv: (args) => {
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
      texthookerEnabled: false,
    }),
    getDefaultMpvLogPath: () => '/tmp/mp.log',
    defaultMpvArgs: ['--sid=auto'],
    removeSocketPath: () => {},
    spawnMpv: (args) => {
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

test('createLaunchMpvIdleForJellyfinPlaybackHandler skips bundled script when installed plugin exists', () => {
  const spawnedArgs: string[][] = [];
  const launch = createLaunchMpvIdleForJellyfinPlaybackHandler({
    getSocketPath: () => '/tmp/subminer.sock',
    getLaunchMode: () => 'normal',
    platform: 'linux',
    execPath: '/opt/SubMiner/SubMiner.AppImage',
    getRuntimePluginEntrypoint: () => '/opt/SubMiner/plugin/subminer/main.lua',
    getInstalledPluginDetection: () => ({
      installed: true,
      path: '/home/tester/.config/mpv/scripts/subminer/main.lua',
      version: '0.1.0',
      source: 'default-config',
      message: null,
    }),
    getDefaultMpvLogPath: () => '/tmp/mp.log',
    defaultMpvArgs: ['--sid=auto'],
    removeSocketPath: () => {},
    spawnMpv: (args) => {
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
  assert.equal(
    spawnedArgs[0]?.some((arg) => arg.startsWith('--script=/opt/SubMiner/plugin/subminer')),
    false,
  );
  assert.ok(spawnedArgs[0]?.some((arg) => arg.startsWith('--script-opts=')));
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

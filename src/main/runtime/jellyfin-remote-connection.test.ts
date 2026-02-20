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
    platform: 'darwin',
    execPath: '/Applications/SubMiner.app/Contents/MacOS/SubMiner',
    defaultMpvLogPath: '/tmp/mp.log',
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
  assert.ok(spawnedArgs[0].includes('--idle=yes'));
  assert.ok(spawnedArgs[0].some((arg) => arg.includes('--input-ipc-server=/tmp/subminer.sock')));
  assert.ok(logs.some((entry) => entry.includes('Launched mpv for Jellyfin playback')));
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

import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildEnsureMpvConnectedForJellyfinPlaybackMainDepsHandler,
  createBuildLaunchMpvIdleForJellyfinPlaybackMainDepsHandler,
  createBuildWaitForMpvConnectedMainDepsHandler,
} from './jellyfin-remote-connection-main-deps';

test('wait for mpv connected main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const client = { connected: false, connect: () => calls.push('connect') };
  const deps = createBuildWaitForMpvConnectedMainDepsHandler({
    getMpvClient: () => client,
    now: () => 123,
    sleep: async () => {
      calls.push('sleep');
    },
  })();

  assert.equal(deps.getMpvClient(), client);
  assert.equal(deps.now(), 123);
  await deps.sleep(10);
  assert.deepEqual(calls, ['sleep']);
});

test('launch mpv for jellyfin main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const proc = {
    on: () => {},
    unref: () => {
      calls.push('unref');
    },
  };
  const deps = createBuildLaunchMpvIdleForJellyfinPlaybackMainDepsHandler({
    getSocketPath: () => '/tmp/mpv.sock',
    getLaunchMode: () => 'fullscreen',
    platform: 'darwin',
    execPath: '/tmp/subminer',
    getRuntimePluginEntrypoint: () => '/tmp/plugin/subminer/main.lua',
    getInstalledPluginDetection: () => ({
      installed: false,
      path: null,
      version: null,
      source: null,
      message: null,
    }),
    defaultMpvLogPath: '/tmp/mpv.log',
    defaultMpvArgs: ['--no-config'],
    removeSocketPath: (socketPath) => calls.push(`rm:${socketPath}`),
    spawnMpv: (args) => {
      calls.push(`spawn:${args.join(' ')}`);
      return proc;
    },
    logWarn: (message) => calls.push(`warn:${message}`),
    logInfo: (message) => calls.push(`info:${message}`),
  })();

  assert.equal(deps.getSocketPath(), '/tmp/mpv.sock');
  assert.equal(deps.getLaunchMode(), 'fullscreen');
  assert.equal(deps.platform, 'darwin');
  assert.equal(deps.execPath, '/tmp/subminer');
  assert.equal(deps.getRuntimePluginEntrypoint?.(), '/tmp/plugin/subminer/main.lua');
  assert.equal(deps.getInstalledPluginDetection?.().installed, false);
  assert.equal(deps.defaultMpvLogPath, '/tmp/mpv.log');
  assert.deepEqual(deps.defaultMpvArgs, ['--no-config']);
  deps.removeSocketPath('/tmp/mpv.sock');
  deps.spawnMpv(['--idle=yes']);
  deps.logInfo('launched');
  deps.logWarn('bad', null);
  assert.deepEqual(calls, ['rm:/tmp/mpv.sock', 'spawn:--idle=yes', 'info:launched', 'warn:bad']);
});

test('ensure mpv connected for jellyfin main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const client = { connected: true, connect: () => {} };
  const waitPromise = Promise.resolve(true);
  const inFlight = Promise.resolve(false);
  const deps = createBuildEnsureMpvConnectedForJellyfinPlaybackMainDepsHandler({
    getMpvClient: () => client,
    setMpvClient: () => calls.push('set-client'),
    createMpvClient: () => client,
    waitForMpvConnected: () => waitPromise,
    launchMpvIdleForJellyfinPlayback: () => calls.push('launch'),
    getAutoLaunchInFlight: () => inFlight,
    setAutoLaunchInFlight: () => calls.push('set-in-flight'),
    connectTimeoutMs: 7000,
    autoLaunchTimeoutMs: 15000,
  })();

  assert.equal(deps.getMpvClient(), client);
  deps.setMpvClient(client);
  assert.equal(deps.createMpvClient(), client);
  assert.equal(await deps.waitForMpvConnected(1), true);
  deps.launchMpvIdleForJellyfinPlayback();
  assert.equal(deps.getAutoLaunchInFlight(), inFlight);
  deps.setAutoLaunchInFlight(null);
  assert.equal(deps.connectTimeoutMs, 7000);
  assert.equal(deps.autoLaunchTimeoutMs, 15000);
  assert.deepEqual(calls, ['set-client', 'launch', 'set-in-flight']);
});

import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createStartJellyfinRemoteSessionHandler,
  createStopJellyfinRemoteSessionHandler,
} from './jellyfin-remote-session-lifecycle';

type StartDeps = Parameters<typeof createStartJellyfinRemoteSessionHandler>[0];
type ServiceOptions = Parameters<StartDeps['createRemoteSessionService']>[0];
type FakeService = ReturnType<StartDeps['createRemoteSessionService']>;

function createConfig(overrides?: Partial<Record<string, unknown>>) {
  return {
    enabled: true,
    remoteControlEnabled: true,
    remoteControlAutoConnect: true,
    serverUrl: 'http://localhost',
    accessToken: 'token',
    userId: 'user-id',
    autoAnnounce: false,
    ...(overrides || {}),
  } as never;
}

function makeService(overrides: Partial<FakeService> = {}): FakeService {
  return {
    start: () => {},
    stop: () => {},
    advertiseNow: async () => true,
    ...overrides,
  };
}

/**
 * Builds the start handler with recording fakes. `hostName` also seeds the client
 * info device id, mirroring the hostname-derived identity used in production.
 */
function makeHarness(
  options: {
    config?: Partial<Record<string, unknown>>;
    hostName?: string;
    existing?: FakeService | null;
    service?: Partial<FakeService>;
    deps?: Partial<StartDeps>;
  } = {},
) {
  const hostName = options.hostName ?? 'workstation';
  const created: ServiceOptions[] = [];
  const infos: string[] = [];
  const warnings: Array<{ message: string; details?: unknown }> = [];
  const state = {
    current: options.existing ?? (null as FakeService | null),
    stateChanges: 0,
    started: 0,
  };

  const startRemote = createStartJellyfinRemoteSessionHandler({
    getJellyfinConfig: () => createConfig(options.config),
    getCurrentSession: () => state.current,
    setCurrentSession: (session) => {
      state.current = session;
    },
    createRemoteSessionService: (serviceOptions) => {
      created.push(serviceOptions);
      return makeService({
        start: () => {
          state.started += 1;
        },
        ...options.service,
      });
    },
    defaultDeviceId: 'default-device',
    defaultClientName: 'SubMiner',
    defaultClientVersion: '1.0',
    getClientInfo: () => ({ deviceId: hostName, clientName: 'SubMiner', clientVersion: '1.0' }),
    getHostName: () => hostName,
    handlePlay: async () => {},
    handlePlaystate: async () => {},
    handleGeneralCommand: async () => {},
    logInfo: (message) => infos.push(message),
    logWarn: (message, details) => warnings.push({ message, details }),
    onSessionStateChanged: () => {
      state.stateChanges += 1;
    },
    ...options.deps,
  });

  return { startRemote, created, infos, warnings, state };
}

const flushPromises = () => new Promise((resolve) => setImmediate(resolve));

const GATING_CASES: Array<{
  name: string;
  config: Partial<Record<string, unknown>>;
  explicitCreates: boolean;
}> = [
  { name: 'jellyfin integration is disabled', config: { enabled: false }, explicitCreates: false },
  {
    name: 'remote control is disabled',
    config: { remoteControlEnabled: false },
    explicitCreates: false,
  },
  {
    name: 'auto-connect is off, unless explicit start is requested',
    config: { remoteControlAutoConnect: false },
    explicitCreates: true,
  },
];

for (const c of GATING_CASES) {
  test(`start handler no-ops when ${c.name}`, async () => {
    const { startRemote, created } = makeHarness({ config: c.config });

    await startRemote();
    assert.equal(created.length, 0);

    await startRemote({ explicit: true });
    assert.equal(created.length, c.explicitCreates ? 1 : 0);
  });
}

test('start handler creates, starts, and stores session', async () => {
  const { startRemote, created, infos, state } = makeHarness();

  await startRemote();

  assert.equal(created.length, 1);
  assert.equal(state.started, 1);
  assert.ok(state.current);
  assert.equal(state.stateChanges, 1);
  assert.ok(infos.some((line) => line.includes('Jellyfin remote session enabled (workstation).')));
});

// The visible device name is always the hostname; a configured name is ignored.
const DEVICE_IDENTITY_CASES: Array<{
  name: string;
  hostName: string;
  config: Partial<Record<string, unknown>>;
}> = [
  {
    name: 'derives client info and device name from the hostname',
    hostName: 'kyle-pc',
    config: {},
  },
  {
    name: 'ignores configured visible device name',
    hostName: 'cachy',
    config: { remoteControlDeviceName: 'SubMiner Cachy sudacode' },
  },
];

for (const c of DEVICE_IDENTITY_CASES) {
  test(`start handler ${c.name}`, async () => {
    const { startRemote, created } = makeHarness({ hostName: c.hostName, config: c.config });

    await startRemote({ explicit: true });

    assert.equal(created.length, 1);
    assert.deepEqual(
      {
        deviceId: created[0]?.deviceId,
        clientName: created[0]?.clientName,
        clientVersion: created[0]?.clientVersion,
        deviceName: created[0]?.deviceName,
      },
      {
        deviceId: c.hostName,
        clientName: 'SubMiner',
        clientVersion: '1.0',
        deviceName: c.hostName,
      },
    );
  });
}

test('start handler stops previous session before replacing', async () => {
  let stopCalls = 0;
  const { startRemote, state } = makeHarness({
    existing: makeService({
      stop: () => {
        stopCalls += 1;
      },
    }),
  });

  await startRemote();

  assert.equal(stopCalls, 1);
  assert.equal(state.started, 1);
});

test('created service announces after connect when autoAnnounce is on', async () => {
  const { startRemote, created, infos, warnings } = makeHarness({ config: { autoAnnounce: true } });
  await startRemote();

  created[0]?.onConnected();
  await flushPromises();

  assert.ok(infos.includes('Jellyfin cast target is visible to server sessions.'));
  assert.deepEqual(warnings, []);
});

test('created service warns when announced device is not visible yet', async () => {
  const { startRemote, created, infos, warnings } = makeHarness({
    config: { autoAnnounce: true },
    service: { advertiseNow: async () => false },
  });
  await startRemote();

  created[0]?.onConnected();
  await flushPromises();

  assert.deepEqual(
    warnings.map((entry) => entry.message),
    ['Jellyfin remote connected but device not visible in server sessions yet.'],
  );
  assert.ok(!infos.includes('Jellyfin cast target is visible to server sessions.'));
});

test('created service does not announce on connect when autoAnnounce is off', async () => {
  let advertised = 0;
  const { startRemote, created } = makeHarness({
    service: {
      advertiseNow: async () => {
        advertised += 1;
        return true;
      },
    },
  });
  await startRemote();

  created[0]?.onConnected();
  await flushPromises();

  assert.equal(advertised, 0);
});

const EVENT_FAILURE_CASES: Array<{
  event: 'Play' | 'Playstate' | 'GeneralCommand';
  handler: 'handlePlay' | 'handlePlaystate' | 'handleGeneralCommand';
  callback: 'onPlay' | 'onPlaystate' | 'onGeneralCommand';
}> = [
  { event: 'Play', handler: 'handlePlay', callback: 'onPlay' },
  { event: 'Playstate', handler: 'handlePlaystate', callback: 'onPlaystate' },
  { event: 'GeneralCommand', handler: 'handleGeneralCommand', callback: 'onGeneralCommand' },
];

for (const c of EVENT_FAILURE_CASES) {
  test(`created service logs a warning when the ${c.event} handler rejects`, async () => {
    const failure = new Error('boom');
    const { startRemote, created, warnings } = makeHarness({
      deps: {
        [c.handler]: async () => {
          throw failure;
        },
      },
    });
    await startRemote();

    created[0]?.[c.callback]({});
    await flushPromises();

    assert.deepEqual(warnings, [
      { message: `Failed handling Jellyfin remote ${c.event} event`, details: failure },
    ]);
  });
}

test('stop handler stops active session and clears playback', () => {
  let stopCalls = 0;
  let clearCalls = 0;
  let stateChanges = 0;
  let currentSession: FakeService | null = makeService({
    stop: () => {
      stopCalls += 1;
    },
  });

  const stopRemote = createStopJellyfinRemoteSessionHandler({
    getCurrentSession: () => currentSession,
    setCurrentSession: (session) => {
      currentSession = session;
    },
    clearActivePlayback: () => {
      clearCalls += 1;
    },
    onSessionStateChanged: () => {
      stateChanges += 1;
    },
  });

  stopRemote();
  assert.equal(stopCalls, 1);
  assert.equal(clearCalls, 1);
  assert.equal(currentSession, null);
  assert.equal(stateChanges, 1);
});

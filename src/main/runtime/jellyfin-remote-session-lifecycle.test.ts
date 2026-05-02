import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createStartJellyfinRemoteSessionHandler,
  createStopJellyfinRemoteSessionHandler,
} from './jellyfin-remote-session-lifecycle';

function createConfig(overrides?: Partial<Record<string, unknown>>) {
  return {
    enabled: true,
    remoteControlEnabled: true,
    remoteControlAutoConnect: true,
    serverUrl: 'http://localhost',
    accessToken: 'token',
    userId: 'user-id',
    deviceId: '',
    clientName: '',
    clientVersion: '',
    remoteControlDeviceName: '',
    autoAnnounce: false,
    ...(overrides || {}),
  } as never;
}

test('start handler no-ops when jellyfin integration is disabled', async () => {
  let created = false;
  const startRemote = createStartJellyfinRemoteSessionHandler({
    getJellyfinConfig: () => createConfig({ enabled: false }),
    getCurrentSession: () => null,
    setCurrentSession: () => {},
    createRemoteSessionService: () => {
      created = true;
      return {
        start: () => {},
        stop: () => {},
        advertiseNow: async () => true,
      };
    },
    defaultDeviceId: 'default-device',
    defaultClientName: 'SubMiner',
    defaultClientVersion: '1.0',
    handlePlay: async () => {},
    handlePlaystate: async () => {},
    handleGeneralCommand: async () => {},
    logInfo: () => {},
    logWarn: () => {},
  });

  await startRemote();
  assert.equal(created, false);
});

test('start handler no-ops when remote control is disabled', async () => {
  let created = false;
  const startRemote = createStartJellyfinRemoteSessionHandler({
    getJellyfinConfig: () => createConfig({ remoteControlEnabled: false }),
    getCurrentSession: () => null,
    setCurrentSession: () => {},
    createRemoteSessionService: () => {
      created = true;
      return {
        start: () => {},
        stop: () => {},
        advertiseNow: async () => true,
      };
    },
    defaultDeviceId: 'default-device',
    defaultClientName: 'SubMiner',
    defaultClientVersion: '1.0',
    handlePlay: async () => {},
    handlePlaystate: async () => {},
    handleGeneralCommand: async () => {},
    logInfo: () => {},
    logWarn: () => {},
  });

  await startRemote();
  assert.equal(created, false);
});

test('start handler respects auto-connect unless explicit start is requested', async () => {
  let created = 0;
  const startRemote = createStartJellyfinRemoteSessionHandler({
    getJellyfinConfig: () => createConfig({ remoteControlAutoConnect: false }),
    getCurrentSession: () => null,
    setCurrentSession: () => {},
    createRemoteSessionService: () => {
      created += 1;
      return {
        start: () => {},
        stop: () => {},
        advertiseNow: async () => true,
      };
    },
    defaultDeviceId: 'default-device',
    defaultClientName: 'SubMiner',
    defaultClientVersion: '1.0',
    handlePlay: async () => {},
    handlePlaystate: async () => {},
    handleGeneralCommand: async () => {},
    logInfo: () => {},
    logWarn: () => {},
  });

  await startRemote();
  assert.equal(created, 0);

  await startRemote({ explicit: true });
  assert.equal(created, 1);
});

test('start handler creates, starts, and stores session', async () => {
  let storedSession: {
    start: () => void;
    stop: () => void;
    advertiseNow: () => Promise<boolean>;
  } | null = null;
  let started = false;
  const infos: string[] = [];
  const startRemote = createStartJellyfinRemoteSessionHandler({
    getJellyfinConfig: () => createConfig({ clientName: 'Desk' }),
    getCurrentSession: () => null,
    setCurrentSession: (session) => {
      storedSession = session as never;
    },
    createRemoteSessionService: (options) => {
      assert.equal(options.deviceName, 'Desk');
      return {
        start: () => {
          started = true;
        },
        stop: () => {},
        advertiseNow: async () => true,
      };
    },
    defaultDeviceId: 'default-device',
    defaultClientName: 'SubMiner',
    defaultClientVersion: '1.0',
    handlePlay: async () => {},
    handlePlaystate: async () => {},
    handleGeneralCommand: async () => {},
    logInfo: (message) => infos.push(message),
    logWarn: () => {},
  });

  await startRemote();

  assert.equal(started, true);
  assert.ok(storedSession);
  assert.ok(infos.some((line) => line.includes('Jellyfin remote session enabled (Desk).')));
});

test('start handler stops previous session before replacing', async () => {
  let stopCalls = 0;
  const oldSession = {
    start: () => {},
    stop: () => {
      stopCalls += 1;
    },
    advertiseNow: async () => true,
  };
  let current: typeof oldSession | null = oldSession;

  const startRemote = createStartJellyfinRemoteSessionHandler({
    getJellyfinConfig: () => createConfig(),
    getCurrentSession: () => current,
    setCurrentSession: (session) => {
      current = session as never;
    },
    createRemoteSessionService: () => ({
      start: () => {},
      stop: () => {},
      advertiseNow: async () => true,
    }),
    defaultDeviceId: 'default-device',
    defaultClientName: 'SubMiner',
    defaultClientVersion: '1.0',
    handlePlay: async () => {},
    handlePlaystate: async () => {},
    handleGeneralCommand: async () => {},
    logInfo: () => {},
    logWarn: () => {},
  });

  await startRemote();
  assert.equal(stopCalls, 1);
});

test('stop handler stops active session and clears playback', () => {
  let stopCalls = 0;
  let clearCalls = 0;
  let currentSession: { stop: () => void } | null = {
    stop: () => {
      stopCalls += 1;
    },
  };

  const stopRemote = createStopJellyfinRemoteSessionHandler({
    getCurrentSession: () => currentSession as never,
    setCurrentSession: (session) => {
      currentSession = session as never;
    },
    clearActivePlayback: () => {
      clearCalls += 1;
    },
  });

  stopRemote();
  assert.equal(stopCalls, 1);
  assert.equal(clearCalls, 1);
  assert.equal(currentSession, null);
});

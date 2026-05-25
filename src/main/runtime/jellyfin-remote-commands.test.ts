import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createHandleJellyfinRemoteGeneralCommand,
  createHandleJellyfinRemotePlay,
  createHandleJellyfinRemotePlaystate,
  getConfiguredJellyfinSession,
  type ActiveJellyfinRemotePlaybackState,
} from './jellyfin-remote-commands';

test('getConfiguredJellyfinSession returns null for incomplete config', () => {
  assert.equal(
    getConfiguredJellyfinSession({
      serverUrl: '',
      accessToken: 'token',
      userId: 'user',
      username: 'name',
    }),
    null,
  );
});

test('createHandleJellyfinRemotePlay forwards parsed payload to play runtime', async () => {
  const calls: Array<{
    itemId: string;
    audio?: number;
    subtitle?: number;
    start?: number;
    fallback?: boolean;
  }> = [];
  const handlePlay = createHandleJellyfinRemotePlay({
    getConfiguredSession: () => ({
      serverUrl: 'https://jellyfin.local',
      accessToken: 'token',
      userId: 'user',
      username: 'name',
    }),
    getClientInfo: () => ({ clientName: 'SubMiner', clientVersion: '1.0', deviceId: 'abc' }),
    getJellyfinConfig: () => ({ enabled: true }),
    playJellyfinItem: async (params) => {
      calls.push({
        itemId: params.itemId,
        audio: params.audioStreamIndex,
        subtitle: params.subtitleStreamIndex,
        start: params.startTimeTicksOverride,
        fallback: params.fallbackToPlanStartTimeOnZeroOverride,
      });
    },
    logWarn: () => {},
  });

  await handlePlay({
    ItemIds: ['item-1'],
    AudioStreamIndex: 3,
    SubtitleStreamIndex: 7,
    StartPositionTicks: 1000,
  });

  assert.deepEqual(calls, [
    { itemId: 'item-1', audio: 3, subtitle: 7, start: 1000, fallback: true },
  ]);
});

test('createHandleJellyfinRemotePlay parses string StartPositionTicks', async () => {
  const calls: Array<{ itemId: string; start?: number; fallback?: boolean }> = [];
  const handlePlay = createHandleJellyfinRemotePlay({
    getConfiguredSession: () => ({
      serverUrl: 'https://jellyfin.local',
      accessToken: 'token',
      userId: 'user',
      username: 'name',
    }),
    getClientInfo: () => ({ clientName: 'SubMiner', clientVersion: '1.0', deviceId: 'abc' }),
    getJellyfinConfig: () => ({ enabled: true }),
    playJellyfinItem: async (params) => {
      calls.push({
        itemId: params.itemId,
        start: params.startTimeTicksOverride,
        fallback: params.fallbackToPlanStartTimeOnZeroOverride,
      });
    },
    logWarn: () => {},
  });

  await handlePlay({
    ItemIds: ['item-2'],
    StartPositionTicks: '12345',
  });

  assert.deepEqual(calls, [{ itemId: 'item-2', start: 12345, fallback: true }]);
});

test('createHandleJellyfinRemotePlay starts from beginning when StartPositionTicks is omitted', async () => {
  const calls: Array<{ itemId: string; start?: number; fallback?: boolean }> = [];
  const handlePlay = createHandleJellyfinRemotePlay({
    getConfiguredSession: () => ({
      serverUrl: 'https://jellyfin.local',
      accessToken: 'token',
      userId: 'user',
      username: 'name',
    }),
    getClientInfo: () => ({ clientName: 'SubMiner', clientVersion: '1.0', deviceId: 'abc' }),
    getJellyfinConfig: () => ({ enabled: true }),
    playJellyfinItem: async (params) => {
      calls.push({
        itemId: params.itemId,
        start: params.startTimeTicksOverride,
        fallback: params.fallbackToPlanStartTimeOnZeroOverride,
      });
    },
    logWarn: () => {},
  });

  await handlePlay({
    ItemIds: ['item-3'],
  });

  assert.deepEqual(calls, [{ itemId: 'item-3', start: 0, fallback: false }]);
});

test('createHandleJellyfinRemotePlay lets explicit zero fall back to Jellyfin item progress', async () => {
  const calls: Array<{ itemId: string; start?: number; fallback?: boolean }> = [];
  const handlePlay = createHandleJellyfinRemotePlay({
    getConfiguredSession: () => ({
      serverUrl: 'https://jellyfin.local',
      accessToken: 'token',
      userId: 'user',
      username: 'name',
    }),
    getClientInfo: () => ({ clientName: 'SubMiner', clientVersion: '1.0', deviceId: 'abc' }),
    getJellyfinConfig: () => ({ enabled: true }),
    playJellyfinItem: async (params) => {
      calls.push({
        itemId: params.itemId,
        start: params.startTimeTicksOverride,
        fallback: params.fallbackToPlanStartTimeOnZeroOverride,
      });
    },
    logWarn: () => {},
  });

  await handlePlay({
    ItemIds: ['item-4'],
    StartPositionTicks: 0,
  });

  assert.deepEqual(calls, [{ itemId: 'item-4', start: 0, fallback: true }]);
});

test('createHandleJellyfinRemotePlay logs and skips payload without item id', async () => {
  const warnings: string[] = [];
  const handlePlay = createHandleJellyfinRemotePlay({
    getConfiguredSession: () => ({
      serverUrl: 'https://jellyfin.local',
      accessToken: 'token',
      userId: 'user',
      username: 'name',
    }),
    getClientInfo: () => ({ clientName: 'SubMiner', clientVersion: '1.0', deviceId: 'abc' }),
    getJellyfinConfig: () => ({}),
    playJellyfinItem: async () => {
      throw new Error('should not be called');
    },
    logWarn: (message) => warnings.push(message),
  });

  await handlePlay({ ItemIds: [] });
  assert.deepEqual(warnings, ['Ignoring Jellyfin remote Play event without ItemIds.']);
});

test('createHandleJellyfinRemotePlay ignores duplicate play for active item', async () => {
  let playCalls = 0;
  const handlePlay = createHandleJellyfinRemotePlay({
    getConfiguredSession: () => ({
      serverUrl: 'https://jellyfin.local',
      accessToken: 'token',
      userId: 'user',
      username: 'name',
    }),
    getClientInfo: () => ({ clientName: 'SubMiner', clientVersion: '1.0', deviceId: 'abc' }),
    getJellyfinConfig: () => ({}),
    getActivePlayback: () => ({
      itemId: 'item-1',
      playMethod: 'DirectPlay',
    }),
    playJellyfinItem: async () => {
      playCalls += 1;
    },
    logWarn: () => {},
  });

  await handlePlay({ ItemIds: ['item-1'] });

  assert.equal(playCalls, 0);
});

test('createHandleJellyfinRemotePlaystate dispatches pause/seek/stop flows', async () => {
  const mpvClient = {};
  const commands: Array<(string | number)[]> = [];
  const calls: string[] = [];
  const handlePlaystate = createHandleJellyfinRemotePlaystate({
    getMpvClient: () => mpvClient,
    sendMpvCommand: (_client, command) => commands.push(command),
    reportJellyfinRemoteProgress: async (force) => {
      calls.push(`progress:${force}`);
    },
    reportJellyfinRemoteStopped: async () => {
      calls.push('stopped');
    },
    jellyfinTicksToSeconds: (ticks) => ticks / 10,
  });

  await handlePlaystate({ Command: 'Pause' });
  await handlePlaystate({ Command: 'Seek', SeekPositionTicks: 50 });
  await handlePlaystate({ Command: 'Stop' });

  assert.deepEqual(commands, [
    ['set_property', 'pause', 'yes'],
    ['seek', 5, 'absolute+exact'],
    ['stop'],
  ]);
  assert.deepEqual(calls, ['progress:true', 'progress:true', 'stopped']);
});

test('createHandleJellyfinRemoteGeneralCommand mutates active playback indices', async () => {
  const mpvClient = {};
  const commands: Array<(string | number)[]> = [];
  const playback: ActiveJellyfinRemotePlaybackState = {
    itemId: 'item-1',
    playMethod: 'DirectPlay',
    audioStreamIndex: null,
    subtitleStreamIndex: null,
  };
  const calls: string[] = [];

  const handleGeneral = createHandleJellyfinRemoteGeneralCommand({
    getMpvClient: () => mpvClient,
    sendMpvCommand: (_client, command) => commands.push(command),
    getActivePlayback: () => playback,
    reportJellyfinRemoteProgress: async (force) => {
      calls.push(`progress:${force}`);
    },
    logDebug: (message) => {
      calls.push(`debug:${message}`);
    },
  });

  await handleGeneral({ Name: 'SetAudioStreamIndex', Arguments: { Index: 2 } });
  await handleGeneral({ Name: 'SetSubtitleStreamIndex', Arguments: { Index: -1 } });
  await handleGeneral({ Name: 'UnsupportedCommand', Arguments: {} });

  assert.deepEqual(commands, [
    ['set_property', 'aid', 2],
    ['set_property', 'sid', 'no'],
  ]);
  assert.equal(playback.audioStreamIndex, 2);
  assert.equal(playback.subtitleStreamIndex, null);
  assert.ok(calls.includes('progress:true'));
  assert.ok(
    calls.some((entry) =>
      entry.includes('Ignoring unsupported Jellyfin GeneralCommand: UnsupportedCommand'),
    ),
  );
});

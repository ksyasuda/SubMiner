import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createHandleJellyfinRemoteGeneralCommand,
  createHandleJellyfinRemotePlay,
  createHandleJellyfinRemotePlaystate,
  getConfiguredJellyfinSession,
  type ActiveJellyfinRemotePlaybackState,
  type JellyfinRemotePlayHandlerDeps,
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

function makePlayHandler(overrides: Partial<JellyfinRemotePlayHandlerDeps> = {}) {
  const played: Array<
    Pick<
      Parameters<JellyfinRemotePlayHandlerDeps['playJellyfinItem']>[0],
      | 'itemId'
      | 'audioStreamIndex'
      | 'subtitleStreamIndex'
      | 'startTimeTicksOverride'
      | 'fallbackToPlanStartTimeOnZeroOverride'
    >
  > = [];
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
    playJellyfinItem: async (params) => {
      played.push({
        itemId: params.itemId,
        audioStreamIndex: params.audioStreamIndex,
        subtitleStreamIndex: params.subtitleStreamIndex,
        startTimeTicksOverride: params.startTimeTicksOverride,
        fallbackToPlanStartTimeOnZeroOverride: params.fallbackToPlanStartTimeOnZeroOverride,
      });
    },
    logWarn: (message) => warnings.push(message),
    ...overrides,
  });
  return { handlePlay, played, warnings };
}

const PLAY_PAYLOAD_CASES: Array<{
  name: string;
  payload: Record<string, unknown>;
  expected: {
    audioStreamIndex?: number;
    subtitleStreamIndex?: number;
    startTimeTicksOverride: number;
    fallbackToPlanStartTimeOnZeroOverride: boolean;
  };
}> = [
  {
    name: 'forwards audio, subtitle and numeric StartPositionTicks',
    payload: { AudioStreamIndex: 3, SubtitleStreamIndex: 7, StartPositionTicks: 1000 },
    expected: {
      audioStreamIndex: 3,
      subtitleStreamIndex: 7,
      startTimeTicksOverride: 1000,
      fallbackToPlanStartTimeOnZeroOverride: true,
    },
  },
  {
    name: 'parses string StartPositionTicks',
    payload: { StartPositionTicks: '12345' },
    expected: { startTimeTicksOverride: 12345, fallbackToPlanStartTimeOnZeroOverride: true },
  },
  {
    name: 'starts from beginning when StartPositionTicks is omitted',
    payload: {},
    expected: { startTimeTicksOverride: 0, fallbackToPlanStartTimeOnZeroOverride: false },
  },
  {
    name: 'lets explicit zero fall back to Jellyfin item progress',
    payload: { StartPositionTicks: 0 },
    expected: { startTimeTicksOverride: 0, fallbackToPlanStartTimeOnZeroOverride: true },
  },
];

for (const c of PLAY_PAYLOAD_CASES) {
  test(`createHandleJellyfinRemotePlay ${c.name}`, async () => {
    const { handlePlay, played } = makePlayHandler();

    await handlePlay({ ItemIds: ['item-1'], ...c.payload });

    assert.deepEqual(played, [
      {
        itemId: 'item-1',
        audioStreamIndex: undefined,
        subtitleStreamIndex: undefined,
        ...c.expected,
      },
    ]);
  });
}

test('createHandleJellyfinRemotePlay logs and skips payload without item id', async () => {
  const { handlePlay, played, warnings } = makePlayHandler();

  await handlePlay({ ItemIds: [] });

  assert.deepEqual(warnings, ['Ignoring Jellyfin remote Play event without ItemIds.']);
  assert.deepEqual(played, []);
});

test('createHandleJellyfinRemotePlay ignores duplicate play for active item', async () => {
  const { handlePlay, played } = makePlayHandler({
    getActivePlayback: () => ({ itemId: 'item-1', playMethod: 'DirectPlay' }),
  });

  await handlePlay({ ItemIds: ['item-1'] });

  assert.deepEqual(played, []);
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

test('createHandleJellyfinRemotePlaystate maps Unpause and PlayPause and reports progress', async () => {
  const commands: Array<(string | number)[]> = [];
  const progress: boolean[] = [];
  const handlePlaystate = createHandleJellyfinRemotePlaystate({
    getMpvClient: () => ({}),
    sendMpvCommand: (_client, command) => commands.push(command),
    reportJellyfinRemoteProgress: async (force) => {
      progress.push(force);
    },
    reportJellyfinRemoteStopped: async () => {},
    jellyfinTicksToSeconds: (ticks) => ticks / 10,
  });

  await handlePlaystate({ Command: 'Unpause' });
  await handlePlaystate({ Command: 'PlayPause' });

  assert.deepEqual(commands, [
    ['set_property', 'pause', 'no'],
    ['cycle', 'pause'],
  ]);
  assert.deepEqual(progress, [true, true]);
});

test('createHandleJellyfinRemotePlaystate ignores Seek without a valid tick position', async () => {
  const commands: Array<(string | number)[]> = [];
  const progress: boolean[] = [];
  const handlePlaystate = createHandleJellyfinRemotePlaystate({
    getMpvClient: () => ({}),
    sendMpvCommand: (_client, command) => commands.push(command),
    reportJellyfinRemoteProgress: async (force) => {
      progress.push(force);
    },
    reportJellyfinRemoteStopped: async () => {},
    jellyfinTicksToSeconds: (ticks) => ticks / 10,
  });

  await handlePlaystate({ Command: 'Seek', SeekPositionTicks: 'soon' });
  await handlePlaystate({ Command: 'Seek' });

  assert.deepEqual(commands, []);
  assert.deepEqual(progress, []);
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

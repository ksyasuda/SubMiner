import assert from 'node:assert/strict';
import test from 'node:test';
import { createSubtitleSelectionRuntime } from './subtitle-selection';

function setup() {
  let enabled = true;
  const properties = new Map<string, unknown>([
    ['path', '/video.mkv'],
    [
      'track-list',
      [
        { id: 1, type: 'audio' },
        { id: 2, type: 'sub', title: 'Japanese', lang: 'ja', codec: 'ass' },
        { id: 3, type: 'sub', title: 'English', lang: 'en', external: true },
        { id: '4', type: 'sub' },
      ],
    ],
    ['sid', 2],
    ['secondary-sid', 3],
  ]);
  const commands: unknown[][] = [];
  const client = {
    connected: true,
    requestProperty: async (name: string) => properties.get(name),
    request: async (command: unknown[]) => {
      commands.push(command);
      return { error: 'success' };
    },
  };
  const runtime = createSubtitleSelectionRuntime({
    isEnabled: () => enabled,
    getMpvClient: () => client,
  });
  return {
    runtime,
    properties,
    commands,
    client,
    disable: () => {
      enabled = false;
    },
  };
}

test('subtitle selector lists only valid subtitle tracks and current selections', async () => {
  const { runtime, properties } = setup();
  assert.deepEqual(await runtime.getState(), {
    mediaPath: '/video.mkv',
    primary: 2,
    secondary: 3,
    tracks: [
      { id: 2, label: '#2 · Japanese · ja · ass' },
      { id: 3, label: '#3 · English · en · external' },
    ],
  });
  properties.set('sid', 'no');
  properties.set('secondary-sid', false);
  const state = await runtime.getState();
  assert.equal(state.primary, null);
  assert.equal(state.secondary, null);
});

test('subtitle selector swaps tracks and supports disabling both tracks', async () => {
  const { runtime, commands } = setup();
  await runtime.apply({ mediaPath: '/video.mkv', primary: 3, secondary: 2 });
  assert.deepEqual(commands, [
    ['set_property', 'secondary-sid', 'no'],
    ['set_property', 'sid', 3],
    ['set_property', 'secondary-sid', 2],
  ]);
  commands.length = 0;
  await runtime.apply({ mediaPath: '/video.mkv', primary: null, secondary: null });
  assert.ok(commands.every((command) => command[2] === 'no'));
});

test('subtitle selector rejects stale media, unavailable tracks, duplicate tracks and malformed requests without mutation', async () => {
  const { runtime, commands } = setup();
  for (const request of [
    { mediaPath: '/other.mkv', primary: 2, secondary: 3 },
    { mediaPath: '/video.mkv', primary: 99, secondary: null },
    { mediaPath: '/video.mkv', primary: 2, secondary: 2 },
    { mediaPath: '/video.mkv', primary: '2', secondary: null },
    { mediaPath: '/video.mkv', primary: -1, secondary: null },
    null,
  ])
    await assert.rejects(runtime.apply(request));
  assert.deepEqual(commands, []);
});

test('subtitle selector gates access on config and connection and propagates mpv failures', async () => {
  const { runtime, client, disable } = setup();
  client.request = async () => ({ error: 'property unavailable' });
  await assert.rejects(
    runtime.apply({ mediaPath: '/video.mkv', primary: 3, secondary: 2 }),
    /property unavailable/,
  );
  client.connected = false;
  await assert.rejects(runtime.getState(), /Connect to mpv/);
  disable();
  await assert.rejects(runtime.getState(), /Enable subtitle selection/);
});

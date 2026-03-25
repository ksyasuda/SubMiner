import test from 'node:test';
import assert from 'node:assert/strict';
import { createPrepareYoutubePlaybackInMpvHandler } from './youtube-playback-launch';

function createWaitStub() {
  return async (_ms: number): Promise<void> => {};
}

test('prepare youtube playback skips load when current path already matches exact URL', async () => {
  const commands: Array<Array<string>> = [];
  const prepare = createPrepareYoutubePlaybackInMpvHandler({
    requestPath: async () => 'https://www.youtube.com/watch?v=abc123',
    requestProperty: async () => [{ type: 'video', id: 1 }],
    sendMpvCommand: (command) => commands.push(command),
    wait: createWaitStub(),
  });

  const ok = await prepare({ url: 'https://www.youtube.com/watch?v=abc123' });

  assert.equal(ok, true);
  assert.deepEqual(commands, []);
});

test('prepare youtube playback treats matching video IDs as already loaded', async () => {
  const commands: Array<Array<string>> = [];
  const prepare = createPrepareYoutubePlaybackInMpvHandler({
    requestPath: async () => 'https://youtu.be/abc123?t=5',
    requestProperty: async () => [{ type: 'video', id: 1 }],
    sendMpvCommand: (command) => commands.push(command),
    wait: createWaitStub(),
  });

  const ok = await prepare({ url: 'https://www.youtube.com/watch?v=abc123' });

  assert.equal(ok, true);
  assert.deepEqual(commands, []);
});

test('prepare youtube playback replaces media and waits for path switch', async () => {
  const commands: Array<Array<string>> = [];
  const observedPaths = [
    '/videos/episode01.mkv',
    '/videos/episode01.mkv',
    'https://www.youtube.com/watch?v=newvid',
  ];
  const observedTrackLists = [null, [], [{ type: 'video', id: 1 }]];
  let requestCount = 0;
  const prepare = createPrepareYoutubePlaybackInMpvHandler({
    requestPath: async () => {
      const value = observedPaths[Math.min(requestCount, observedPaths.length - 1)] ?? null;
      requestCount += 1;
      return value;
    },
    requestProperty: async (name) => {
      if (name !== 'track-list') return null;
      return observedTrackLists[Math.min(requestCount - 1, observedTrackLists.length - 1)] ?? [];
    },
    sendMpvCommand: (command) => commands.push(command),
    wait: createWaitStub(),
  });

  const ok = await prepare({
    url: 'https://www.youtube.com/watch?v=newvid',
    timeoutMs: 1500,
    pollIntervalMs: 1,
  });

  assert.equal(ok, true);
  assert.deepEqual(commands, [
    ['set_property', 'pause', 'yes'],
    ['set_property', 'sub-auto', 'no'],
    ['set_property', 'sid', 'no'],
    ['set_property', 'secondary-sid', 'no'],
    ['loadfile', 'https://www.youtube.com/watch?v=newvid', 'replace'],
  ]);
});

test('prepare youtube playback returns false after timeout when path never updates', async () => {
  const commands: Array<Array<string>> = [];
  let nowTick = 0;
  const prepare = createPrepareYoutubePlaybackInMpvHandler({
    requestPath: async () => '/videos/episode01.mkv',
    requestProperty: async () => [],
    sendMpvCommand: (command) => commands.push(command),
    wait: createWaitStub(),
    now: () => {
      nowTick += 100;
      return nowTick;
    },
  });

  const ok = await prepare({
    url: 'https://www.youtube.com/watch?v=never-switches',
    timeoutMs: 350,
    pollIntervalMs: 1,
  });

  assert.equal(ok, false);
  assert.deepEqual(commands[4], [
    'loadfile',
    'https://www.youtube.com/watch?v=never-switches',
    'replace',
  ]);
});

test('prepare youtube playback waits for playable media tracks after youtube path matches', async () => {
  const commands: Array<Array<string>> = [];
  const observedPaths = [
    '/videos/episode01.mkv',
    'https://www.youtube.com/watch?v=newvid',
    'https://www.youtube.com/watch?v=newvid',
  ];
  const observedTrackLists = [[], [], [{ type: 'audio', id: 1 }]];
  let requestCount = 0;
  const prepare = createPrepareYoutubePlaybackInMpvHandler({
    requestPath: async () => {
      const value = observedPaths[Math.min(requestCount, observedPaths.length - 1)] ?? null;
      requestCount += 1;
      return value;
    },
    requestProperty: async (name) => {
      if (name !== 'track-list') return null;
      return observedTrackLists[Math.min(requestCount - 1, observedTrackLists.length - 1)] ?? [];
    },
    sendMpvCommand: (command) => commands.push(command),
    wait: createWaitStub(),
  });

  const ok = await prepare({
    url: 'https://www.youtube.com/watch?v=newvid',
    timeoutMs: 1500,
    pollIntervalMs: 1,
  });

  assert.equal(ok, true);
  assert.deepEqual(commands[4], ['loadfile', 'https://www.youtube.com/watch?v=newvid', 'replace']);
});

test('prepare youtube playback accepts a non-youtube resolved path once playable tracks exist', async () => {
  const commands: Array<Array<string>> = [];
  const observedPaths = [
    '/videos/episode01.mkv',
    'https://rr16---sn.example.googlevideo.com/videoplayback?id=abc',
  ];
  const observedTrackLists = [[], [{ type: 'video', id: 1 }, { type: 'audio', id: 2 }]];
  let requestCount = 0;
  const prepare = createPrepareYoutubePlaybackInMpvHandler({
    requestPath: async () => {
      const value = observedPaths[Math.min(requestCount, observedPaths.length - 1)] ?? null;
      requestCount += 1;
      return value;
    },
    requestProperty: async (name) => {
      if (name !== 'track-list') return null;
      return observedTrackLists[Math.min(requestCount - 1, observedTrackLists.length - 1)] ?? [];
    },
    sendMpvCommand: (command) => commands.push(command),
    wait: createWaitStub(),
  });

  const ok = await prepare({
    url: 'https://www.youtube.com/watch?v=newvid',
    timeoutMs: 1500,
    pollIntervalMs: 1,
  });

  assert.equal(ok, true);
  assert.deepEqual(commands[4], ['loadfile', 'https://www.youtube.com/watch?v=newvid', 'replace']);
});

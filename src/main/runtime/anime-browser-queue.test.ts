import test from 'node:test';
import assert from 'node:assert/strict';
import type { AnimeBrowserPlayRequest, AnimeBrowserQueueState } from '../../types/anime-browser';
import type {
  PreparedAnimeBrowserPlayback,
  PrepareAnimeBrowserPlaybackResult,
} from './anime-browser-playback';
import { createAnimeBrowserQueue, type AnimeBrowserQueueDeps } from './anime-browser-queue';

function request(overrides: Partial<AnimeBrowserPlayRequest> = {}): AnimeBrowserPlayRequest {
  return {
    sourceId: 'source',
    animeUrl: '/anime',
    animeTitle: 'Anime',
    episodeUrl: '/episode-1',
    episodeName: 'Episode 1',
    episodeNumber: 1,
    ...overrides,
  };
}

function prepared(input: AnimeBrowserPlayRequest): PreparedAnimeBrowserPlayback {
  const mediaPath = `https://stream.example${input.episodeUrl}.m3u8`;
  return {
    request: input,
    stream: { url: mediaPath, quality: '1080p', headers: {}, audios: [], subtitles: [] },
    metadata: {
      sourceId: input.sourceId,
      animeUrl: input.animeUrl,
      episodeUrl: input.episodeUrl,
      mediaPath,
      statsPath: `animebrowser://${encodeURIComponent(input.episodeUrl)}`,
      seriesTitle: input.animeTitle,
      seasonNumber: null,
      episodeNumber: input.episodeNumber,
      episodeTitle: null,
      displayTitle: `${input.animeTitle} - ${input.episodeName}`,
    },
    trackPreparation: Promise.resolve({
      stream: {
        url: mediaPath,
        quality: '1080p',
        headers: {},
        audios: [],
        subtitles: [],
      },
      subtitleCacheDir: null,
    }),
  };
}

function successful(input: AnimeBrowserPlayRequest): PrepareAnimeBrowserPlaybackResult {
  return { ok: true, playback: prepared(input), error: null, quality: '1080p' };
}

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => {
    resolve = done;
  });
  return { promise, resolve };
}

interface Harness {
  deps: AnimeBrowserQueueDeps;
  prepared: AnimeBrowserPlayRequest[];
  appended: PreparedAnimeBrowserPlayback[];
  activated: PreparedAnimeBrowserPlayback[];
  discarded: PreparedAnimeBrowserPlayback[];
  commands: Array<Array<string | number>>;
  states: AnimeBrowserQueueState[];
  pathChange: (path: string) => void;
  playlist: Array<{ id: number; filename: string; current?: boolean }>;
}

function harness(
  prepareEpisode: AnimeBrowserQueueDeps['prepareEpisode'] = async (input) => successful(input),
): Harness {
  const preparedRequests: AnimeBrowserPlayRequest[] = [];
  const appended: PreparedAnimeBrowserPlayback[] = [];
  const activated: PreparedAnimeBrowserPlayback[] = [];
  const discarded: PreparedAnimeBrowserPlayback[] = [];
  const commands: Array<Array<string | number>> = [];
  const states: AnimeBrowserQueueState[] = [];
  const listeners = new Set<(path: string) => void>();
  const playlist: Array<{ id: number; filename: string; current?: boolean }> = [
    { id: 1, filename: '/current.mkv', current: true },
  ];

  return {
    deps: {
      prepareEpisode: async (input) => {
        preparedRequests.push(input);
        return await prepareEpisode(input);
      },
      appendEpisode: (playback) => {
        appended.push(playback);
        playlist.push({ id: playlist.length + 1, filename: playback.stream.url });
      },
      activateEpisode: async (playback) => void activated.push(playback),
      discardEpisode: async (playback) => void discarded.push(playback),
      armNextEpisode: () => undefined,
      onPlaybackPathChange: (listener) => {
        listeners.add(listener);
        return () => listeners.delete(listener);
      },
      readMpvProperty: async (name) => {
        assert.equal(name, 'playlist');
        return playlist;
      },
      sendMpvCommand: (command) => {
        commands.push(command);
        if (command[0] === 'playlist-remove' && typeof command[1] === 'number') {
          playlist.splice(command[1], 1);
        }
      },
      onQueueState: (state) => void states.push(state),
      log: () => undefined,
    },
    prepared: preparedRequests,
    appended,
    activated,
    discarded,
    commands,
    states,
    pathChange: (path) => {
      for (const listener of [...listeners]) listener(path);
    },
    playlist,
  };
}

test('queue resolves immediately and appends the playable stream to mpv', async () => {
  const h = harness();
  const queue = createAnimeBrowserQueue(h.deps);

  const state = await queue.enqueue(request());

  assert.deepEqual(
    h.prepared.map((entry) => entry.episodeUrl),
    ['/episode-1'],
  );
  assert.deepEqual(
    h.appended.map((entry) => entry.stream.url),
    ['https://stream.example/episode-1.m3u8'],
  );
  assert.deepEqual(
    state.entries.map((entry) => entry.episodeUrl),
    ['/episode-1'],
  );
});

test('concurrent resolutions append in click order', async () => {
  const first = deferred<PrepareAnimeBrowserPlaybackResult>();
  const second = deferred<PrepareAnimeBrowserPlaybackResult>();
  const h = harness((input) =>
    input.episodeUrl === '/episode-1' ? first.promise : second.promise,
  );
  const queue = createAnimeBrowserQueue(h.deps);

  const one = queue.enqueue(request());
  const twoRequest = request({ episodeUrl: '/episode-2', episodeName: 'Episode 2' });
  const two = queue.enqueue(twoRequest);
  second.resolve(successful(twoRequest));
  await new Promise(setImmediate);
  assert.equal(h.appended.length, 0);

  first.resolve(successful(request()));
  await Promise.all([one, two]);
  assert.deepEqual(
    h.appended.map((entry) => entry.request.episodeUrl),
    ['/episode-1', '/episode-2'],
  );
});

test('mpv path navigation advances queue state and activates prepared tracks', async () => {
  const h = harness();
  const queue = createAnimeBrowserQueue(h.deps);
  await queue.enqueue(request());
  await queue.enqueue(request({ episodeUrl: '/episode-2', episodeName: 'Episode 2' }));

  h.pathChange('https://stream.example/episode-1.m3u8');
  await new Promise(setImmediate);

  assert.deepEqual(
    h.activated.map((entry) => entry.request.episodeUrl),
    ['/episode-1'],
  );
  assert.deepEqual(
    queue.getState().entries.map((entry) => entry.episodeUrl),
    ['/episode-2'],
  );
  assert.equal(queue.getState().advances, 1);
  assert.equal(queue.getState().lastStarted?.episodeUrl, '/episode-1');
});

test('dequeue removes the resolved entry from mpv by playlist id', async () => {
  const h = harness();
  const queue = createAnimeBrowserQueue(h.deps);
  await queue.enqueue(request());

  const state = await queue.dequeue('source', '/episode-1');

  assert.deepEqual(state.entries, []);
  assert.deepEqual(h.commands, [['playlist-remove', 1]]);
  assert.equal(h.discarded.length, 1);
});

test('dequeue does not stop an item that mpv began before the request landed', async () => {
  const h = harness();
  const queue = createAnimeBrowserQueue(h.deps);
  await queue.enqueue(request());
  h.playlist[0]!.current = false;
  h.playlist[1]!.current = true;

  await queue.dequeue('source', '/episode-1');

  assert.deepEqual(h.commands, []);
  assert.equal(h.activated.length, 1);
  assert.equal(h.discarded.length, 0);
});

test('clear removes owned playlist entries from the end toward the current file', async () => {
  const h = harness();
  const queue = createAnimeBrowserQueue(h.deps);
  await queue.enqueue(request());
  await queue.enqueue(request({ episodeUrl: '/episode-2', episodeName: 'Episode 2' }));

  await queue.clear();

  assert.deepEqual(h.commands, [
    ['playlist-remove', 2],
    ['playlist-remove', 1],
  ]);
  assert.deepEqual(queue.getState().entries, []);
  assert.equal(h.discarded.length, 2);
});

test('a preparation failure leaves later episodes queued and reports the source error', async () => {
  const h = harness(async (input) =>
    input.episodeUrl === '/episode-1'
      ? { ok: false, playback: null, error: 'No playable video.', quality: null }
      : successful(input),
  );
  const queue = createAnimeBrowserQueue(h.deps);

  const failed = await queue.enqueue(request());
  await queue.enqueue(request({ episodeUrl: '/episode-2', episodeName: 'Episode 2' }));

  assert.equal(failed.lastError, 'Episode 1: No playable video.');
  assert.deepEqual(
    queue.getState().entries.map((entry) => entry.episodeUrl),
    ['/episode-2'],
  );
});

test('queueing the same episode twice resolves and appends it once', async () => {
  const h = harness();
  const queue = createAnimeBrowserQueue(h.deps);

  await queue.enqueue(request());
  await queue.enqueue(request());

  assert.equal(h.prepared.length, 1);
  assert.equal(h.appended.length, 1);
});

test('dequeue while resolution is pending prevents a late append and releases its cache', async () => {
  const pending = deferred<PrepareAnimeBrowserPlaybackResult>();
  const h = harness(() => pending.promise);
  const queue = createAnimeBrowserQueue(h.deps);

  const enqueue = queue.enqueue(request());
  await new Promise(setImmediate);
  await queue.dequeue('source', '/episode-1');
  pending.resolve(successful(request()));
  await enqueue;

  assert.deepEqual(h.appended, []);
  assert.equal(h.discarded.length, 1);
});

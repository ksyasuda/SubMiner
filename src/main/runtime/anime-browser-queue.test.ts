import test from 'node:test';
import assert from 'node:assert/strict';
import type { PlaybackEndFileEvent } from '../../anime-bridge/playback-outcome';
import type { AnimeBrowserPlayRequest, AnimeBrowserQueueState } from '../../types/anime-browser';
import { createAnimeBrowserQueue, type AnimeBrowserQueueDeps } from './anime-browser-queue';

function makeRequest(overrides: Partial<AnimeBrowserPlayRequest> = {}): AnimeBrowserPlayRequest {
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

interface Harness {
  deps: AnimeBrowserQueueDeps;
  played: AnimeBrowserPlayRequest[];
  commands: Array<Array<string | number>>;
  states: AnimeBrowserQueueState[];
  osd: string[];
  endFile: (event: PlaybackEndFileEvent) => void;
  /** How many listeners are subscribed right now. */
  listenerCount: () => number;
}

function makeHarness(
  options: {
    play?: (request: AnimeBrowserPlayRequest) => Promise<{
      ok: boolean;
      error: string | null;
      quality: string | null;
    }>;
    keepOpen?: unknown;
    withEndFile?: boolean;
  } = {},
): Harness {
  const played: AnimeBrowserPlayRequest[] = [];
  const commands: Array<Array<string | number>> = [];
  const states: AnimeBrowserQueueState[] = [];
  const osd: string[] = [];
  const listeners = new Set<(event: PlaybackEndFileEvent) => void>();

  const deps: AnimeBrowserQueueDeps = {
    play: async (request) => {
      played.push(request);
      return options.play
        ? await options.play(request)
        : { ok: true, error: null, quality: '1080p' };
    },
    sendMpvCommand: (command) => void commands.push(command),
    onQueueState: (state) => void states.push(state),
    showMpvOsd: (message) => void osd.push(message),
    log: () => undefined,
  };
  if (options.withEndFile !== false) {
    deps.onPlaybackEndFile = (listener) => {
      listeners.add(listener);
      return () => listeners.delete(listener);
    };
  }
  if (options.keepOpen !== undefined) {
    deps.readMpvProperty = async (name) => {
      if (name !== 'keep-open') throw new Error(`unexpected property ${name}`);
      return options.keepOpen;
    };
  }

  return {
    deps,
    played,
    commands,
    states,
    osd,
    endFile: (event) => {
      for (const listener of [...listeners]) listener(event);
    },
    listenerCount: () => listeners.size,
  };
}

const EOF: PlaybackEndFileEvent = { reason: 'eof', fileError: null };

test('an episode that runs to its end hands the queue its turn', async () => {
  const harness = makeHarness();
  const queue = createAnimeBrowserQueue(harness.deps);

  queue.enqueue(makeRequest());
  queue.enqueue(makeRequest({ episodeUrl: '/episode-2', episodeName: 'Episode 2' }));
  harness.endFile(EOF);
  await new Promise(setImmediate);

  assert.deepEqual(
    harness.played.map((request) => request.episodeUrl),
    ['/episode-1'],
  );
  assert.deepEqual(
    queue.getState().entries.map((entry) => entry.episodeUrl),
    ['/episode-2'],
  );
});

test('only a file that ended by itself advances the queue', async () => {
  const harness = makeHarness();
  const queue = createAnimeBrowserQueue(harness.deps);

  queue.enqueue(makeRequest());
  harness.endFile({ reason: 'stop', fileError: null });
  harness.endFile({ reason: 'quit', fileError: null });
  harness.endFile({ reason: 'error', fileError: 'dead host' });
  await new Promise(setImmediate);

  assert.deepEqual(harness.played, []);
  assert.equal(queue.getState().entries.length, 1);
});

test('queueing the same episode twice leaves one entry', () => {
  const harness = makeHarness();
  const queue = createAnimeBrowserQueue(harness.deps);

  queue.enqueue(makeRequest());
  const state = queue.enqueue(makeRequest());

  assert.equal(state.entries.length, 1);
});

test('an advance is counted and names the episode it started', async () => {
  const harness = makeHarness();
  const queue = createAnimeBrowserQueue(harness.deps);

  queue.enqueue(makeRequest());
  assert.equal(queue.getState().advances, 0);
  harness.endFile(EOF);
  await new Promise(setImmediate);

  const state = queue.getState();
  assert.equal(state.advances, 1);
  assert.equal(state.lastStarted?.episodeUrl, '/episode-1');
});

test('an episode dequeues by its own source and url', () => {
  const harness = makeHarness();
  const queue = createAnimeBrowserQueue(harness.deps);

  queue.enqueue(makeRequest());
  queue.enqueue(makeRequest({ sourceId: 'other' }));
  const state = queue.dequeue('source', '/episode-1');

  assert.deepEqual(
    state.entries.map((entry) => entry.sourceId),
    ['other'],
  );
});

test('a queued episode that will not play stops the queue and reports why', async () => {
  const harness = makeHarness({
    play: async () => ({
      ok: false,
      error: 'That source returned no playable video.',
      quality: null,
    }),
  });
  const queue = createAnimeBrowserQueue(harness.deps);

  queue.enqueue(makeRequest());
  queue.enqueue(makeRequest({ episodeUrl: '/episode-2', episodeName: 'Episode 2' }));
  harness.endFile(EOF);
  await new Promise(setImmediate);

  const state = queue.getState();
  assert.equal(state.lastError, 'Episode 1: That source returned no playable video.');
  // The failed episode is gone; the rest is still the user's queue.
  assert.deepEqual(
    state.entries.map((entry) => entry.episodeUrl),
    ['/episode-2'],
  );
  assert.equal(harness.osd.length, 1);
});

test('a second end-file while an advance is resolving does not double-load', async () => {
  let release = (): void => undefined;
  const started = new Promise<void>((resolve) => {
    release = resolve;
  });
  const harness = makeHarness({
    play: async () => {
      await started;
      return { ok: true, error: null, quality: null };
    },
  });
  const queue = createAnimeBrowserQueue(harness.deps);

  queue.enqueue(makeRequest());
  queue.enqueue(makeRequest({ episodeUrl: '/episode-2', episodeName: 'Episode 2' }));
  harness.endFile(EOF);
  harness.endFile(EOF);
  release();
  await new Promise(setImmediate);

  assert.deepEqual(
    harness.played.map((request) => request.episodeUrl),
    ['/episode-1'],
  );
});

test('mpv is told not to hold the last file open while the queue waits', async () => {
  const harness = makeHarness({ keepOpen: 'yes' });
  const queue = createAnimeBrowserQueue(harness.deps);

  queue.enqueue(makeRequest());
  await new Promise(setImmediate);
  assert.deepEqual(harness.commands, [['set_property', 'keep-open', 'no']]);

  queue.clear();
  assert.deepEqual(harness.commands[1], ['set_property', 'keep-open', 'yes']);
});

test('mpv that already lets files end is left alone', async () => {
  const harness = makeHarness({ keepOpen: 'no' });
  const queue = createAnimeBrowserQueue(harness.deps);

  queue.enqueue(makeRequest());
  await new Promise(setImmediate);
  queue.clear();

  assert.deepEqual(harness.commands, []);
});

test('the last queued episode plays under the user own keep-open setting', async () => {
  const harness = makeHarness({ keepOpen: 'always' });
  const queue = createAnimeBrowserQueue(harness.deps);

  queue.enqueue(makeRequest());
  await new Promise(setImmediate);
  harness.endFile(EOF);
  await new Promise(setImmediate);

  assert.deepEqual(harness.commands, [
    ['set_property', 'keep-open', 'no'],
    ['set_property', 'keep-open', 'always'],
  ]);
});

test('playback started outside the queue re-points the end-file listener', async () => {
  const harness = makeHarness();
  const queue = createAnimeBrowserQueue(harness.deps);

  queue.enqueue(makeRequest());
  queue.handlePlaybackStarted();
  queue.handlePlaybackStarted();

  // Each re-arm replaces the previous subscription rather than stacking on it,
  // so one end-file cannot advance the queue several times over.
  assert.equal(harness.listenerCount(), 1);
  harness.endFile(EOF);
  await new Promise(setImmediate);
  assert.equal(harness.played.length, 1);
});

test('an emptied queue stops listening for the end of the file', () => {
  const harness = makeHarness();
  const queue = createAnimeBrowserQueue(harness.deps);

  queue.enqueue(makeRequest());
  assert.equal(harness.listenerCount(), 1);
  queue.dequeue('source', '/episode-1');
  assert.equal(harness.listenerCount(), 0);
});

test('the queue still accepts episodes when mpv reports no events', () => {
  const harness = makeHarness({ withEndFile: false });
  const queue = createAnimeBrowserQueue(harness.deps);

  assert.equal(queue.enqueue(makeRequest()).entries.length, 1);
});

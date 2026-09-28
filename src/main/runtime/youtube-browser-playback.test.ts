import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createYoutubeBrowserPlaybackRuntime,
  parseYoutubeBrowserVideoRequest,
  type YoutubeBrowserPlaybackDeps,
} from './youtube-browser-playback';

const VIDEO_A = 'https://www.youtube.com/watch?v=aaaaaaaaaaa';
const VIDEO_B = 'https://www.youtube.com/watch?v=bbbbbbbbbbb';

function createHarness(overrides: Partial<YoutubeBrowserPlaybackDeps> = {}) {
  const calls: string[] = [];
  const pendingFlows: Array<() => void> = [];
  let mpvPlaying = false;
  const runtime = createYoutubeBrowserPlaybackRuntime({
    isMpvPlaying: () => mpvPlaying,
    ensureMpvReady: async () => true,
    runPlaybackFlow: (url) => {
      calls.push(`flow:${url}`);
      return new Promise<void>((resolve) => pendingFlows.push(resolve));
    },
    appendToMpvPlaylist: (url) => calls.push(`append:${url}`),
    notifyFailure: (message) => calls.push(`notify:${message}`),
    logWarn: () => {},
    ...overrides,
  });
  return {
    runtime,
    calls,
    setMpvPlaying: (next: boolean) => {
      mpvPlaying = next;
    },
    finishFlows: async () => {
      for (const resolve of pendingFlows.splice(0)) resolve();
      await new Promise((resolve) => setImmediate(resolve));
    },
  };
}

test('parseYoutubeBrowserVideoRequest accepts play and queue requests only', () => {
  assert.deepEqual(parseYoutubeBrowserVideoRequest({ action: 'queue', url: VIDEO_A }), {
    action: 'queue',
    url: VIDEO_A,
  });
  assert.equal(parseYoutubeBrowserVideoRequest({ action: 'delete', url: VIDEO_A }), null);
  assert.equal(parseYoutubeBrowserVideoRequest({ action: 'play' }), null);
  assert.equal(parseYoutubeBrowserVideoRequest(null), null);
});

test('play runs the playback flow with a single-video url', async () => {
  const { runtime, calls } = createHarness();
  const result = await runtime.openVideo({
    action: 'play',
    url: `${VIDEO_A}&list=WL&index=2`,
  });
  assert.deepEqual(result, { ok: true, message: 'Opening in mpv' });
  assert.deepEqual(calls, [`flow:${VIDEO_A}`]);
});

test('queue appends to mpv while something is playing and plays otherwise', async () => {
  const { runtime, calls, setMpvPlaying } = createHarness();
  await runtime.openVideo({ action: 'queue', url: VIDEO_A });
  setMpvPlaying(true);
  const queued = await runtime.openVideo({ action: 'queue', url: VIDEO_B });
  assert.deepEqual(queued, { ok: true, message: 'Queued in mpv' });
  assert.deepEqual(calls, [`flow:${VIDEO_A}`, `append:${VIDEO_B}`]);
});

test('repeat play clicks while mpv is starting open the video once', async () => {
  let releaseMpv: (ready: boolean) => void = () => {};
  const { runtime, calls } = createHarness({
    ensureMpvReady: () =>
      new Promise<boolean>((resolve) => {
        releaseMpv = resolve;
      }),
  });
  const first = runtime.openVideo({ action: 'play', url: VIDEO_A });
  const second = await runtime.openVideo({ action: 'play', url: VIDEO_A });
  releaseMpv(true);
  await first;
  assert.deepEqual(second, { ok: true, message: 'Already opening in mpv' });
  assert.deepEqual(calls, [`flow:${VIDEO_A}`]);
});

test('rejects non-video links and reports when mpv cannot start', async () => {
  const { runtime, calls } = createHarness({ ensureMpvReady: async () => false });
  assert.equal(
    (await runtime.openVideo({ action: 'play', url: 'https://www.youtube.com/' })).ok,
    false,
  );
  assert.deepEqual(await runtime.openVideo({ action: 'play', url: VIDEO_A }), {
    ok: false,
    message: 'Could not start mpv.',
  });
  assert.deepEqual(calls, []);
});

test('advancing onto a queued video runs the playback flow for it', async () => {
  const { runtime, calls, setMpvPlaying, finishFlows } = createHarness();
  await runtime.openVideo({ action: 'play', url: VIDEO_A });
  runtime.handleMediaPathChange(VIDEO_A);
  await finishFlows();
  setMpvPlaying(true);
  await runtime.openVideo({ action: 'queue', url: VIDEO_B });

  runtime.handleMediaPathChange('');
  runtime.handleMediaPathChange(VIDEO_B);
  runtime.handleMediaPathChange(VIDEO_B);

  assert.deepEqual(calls, [`flow:${VIDEO_A}`, `append:${VIDEO_B}`, `flow:${VIDEO_B}`]);
});

test('path changes are ignored while a flow is loading its own video', async () => {
  const { runtime, calls } = createHarness();
  await runtime.openVideo({ action: 'play', url: VIDEO_A });
  await runtime.openVideo({ action: 'play', url: VIDEO_B });
  runtime.handleMediaPathChange(VIDEO_A);
  runtime.handleMediaPathChange(VIDEO_B);
  assert.deepEqual(calls, [`flow:${VIDEO_A}`, `flow:${VIDEO_B}`]);
});

test('a stuck flow does not block queued videos from getting theirs', async () => {
  const { runtime, calls, setMpvPlaying } = createHarness();
  await runtime.openVideo({ action: 'play', url: VIDEO_A });
  runtime.handleMediaPathChange(VIDEO_A);
  setMpvPlaying(true);
  await runtime.openVideo({ action: 'queue', url: VIDEO_B });
  runtime.handleMediaPathChange(VIDEO_B);
  assert.deepEqual(calls, [`flow:${VIDEO_A}`, `append:${VIDEO_B}`, `flow:${VIDEO_B}`]);
});

test('videos not sent from the browser never trigger the flow', async () => {
  const { runtime, calls } = createHarness();
  runtime.handleMediaPathChange(VIDEO_A);
  assert.deepEqual(calls, []);
});

test('mpv disconnect forgets queued videos', async () => {
  const { runtime, calls, setMpvPlaying, finishFlows } = createHarness();
  setMpvPlaying(true);
  await runtime.openVideo({ action: 'queue', url: VIDEO_B });
  runtime.handleMpvDisconnected();
  runtime.handleMediaPathChange(VIDEO_B);
  await finishFlows();
  assert.deepEqual(calls, [`append:${VIDEO_B}`]);
});

test('flow failures are reported', async () => {
  const { runtime, calls } = createHarness({
    runPlaybackFlow: async () => {
      throw new Error('mpv went away');
    },
  });
  await runtime.openVideo({ action: 'play', url: VIDEO_A });
  await new Promise((resolve) => setImmediate(resolve));
  assert.deepEqual(calls, ['notify:YouTube playback failed: mpv went away']);
});

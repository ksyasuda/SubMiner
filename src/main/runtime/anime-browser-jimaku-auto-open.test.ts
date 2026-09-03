import assert from 'node:assert/strict';
import test from 'node:test';
import { createAnimeBrowserJimakuAutoOpen } from './anime-browser-jimaku-auto-open';

function createHarness(options?: {
  enabled?: boolean;
  animeMedia?: boolean | ((mediaPath: string) => boolean);
  paused?: boolean | null;
  opened?: boolean;
  windowReady?: () => Promise<boolean>;
}) {
  const calls: string[] = [];
  let currentMediaPath: string | null = null;
  const runtime = createAnimeBrowserJimakuAutoOpen({
    isEnabled: () => options?.enabled ?? true,
    isAnimeBrowserMedia: (mediaPath) =>
      typeof options?.animeMedia === 'function'
        ? options.animeMedia(mediaPath)
        : (options?.animeMedia ?? true),
    getCurrentMediaPath: () => currentMediaPath,
    getPlaybackPaused: async () => (options?.paused === undefined ? false : options.paused),
    setPlaybackPaused: (paused) => calls.push(`pause:${paused}`),
    waitForPlaybackWindow: () => {
      calls.push('wait-window');
      return options?.windowReady?.() ?? Promise.resolve(true);
    },
    closeAnimeBrowserModal: () => calls.push('close-anime-browser'),
    hideAnimeBrowserWindow: () => calls.push('hide-anime-browser-window'),
    openJimakuModal: async () => {
      calls.push('open-jimaku');
      return options?.opened ?? true;
    },
    logWarn: (message) => calls.push(`warn:${message}`),
  });

  return {
    calls,
    runtime,
    setMediaPath: (mediaPath: string | null) => {
      currentMediaPath = mediaPath;
    },
  };
}

const OPEN_SEQUENCE = [
  'pause:true',
  'wait-window',
  'close-anime-browser',
  'hide-anime-browser-window',
  'open-jimaku',
];

test('anime browser playback pauses, waits for the mpv window, opens Jimaku, and resumes after subtitle load', async () => {
  const harness = createHarness();
  harness.setMediaPath('https://127.0.0.1/stream.m3u8');

  await harness.runtime.handleMediaPathChange('https://127.0.0.1/stream.m3u8');
  assert.deepEqual(harness.calls, OPEN_SEQUENCE);

  harness.runtime.handleJimakuSubtitleLoaded();
  assert.deepEqual(harness.calls, [...OPEN_SEQUENCE, 'pause:false']);
});

test('anime browser playback that was already paused stays paused after subtitle load', async () => {
  const harness = createHarness({ paused: true });
  harness.setMediaPath('https://127.0.0.1/stream.m3u8');

  await harness.runtime.handleMediaPathChange('https://127.0.0.1/stream.m3u8');
  harness.runtime.handleJimakuSubtitleLoaded();

  assert.deepEqual(harness.calls, [
    'wait-window',
    'close-anime-browser',
    'hide-anime-browser-window',
    'open-jimaku',
  ]);
});

test('disabled and non-Anime Browser media do not open Jimaku', async () => {
  const disabled = createHarness({ enabled: false });
  disabled.setMediaPath('/video.mkv');
  await disabled.runtime.handleMediaPathChange('/video.mkv');
  assert.deepEqual(disabled.calls, []);

  const unrelated = createHarness({ animeMedia: false });
  unrelated.setMediaPath('/video.mkv');
  await unrelated.runtime.handleMediaPathChange('/video.mkv');
  assert.deepEqual(unrelated.calls, []);
});

test('a stream that never shows a window releases the pause without opening Jimaku', async () => {
  const harness = createHarness({ windowReady: async () => false });
  harness.setMediaPath('https://127.0.0.1/stream.m3u8');

  await harness.runtime.handleMediaPathChange('https://127.0.0.1/stream.m3u8');

  assert.ok(!harness.calls.includes('open-jimaku'));
  assert.equal(harness.calls.at(-1), 'pause:false');
});

test('a newer episode during the window wait cancels the stale Jimaku open', async () => {
  const firstWait: { resolve: (ready: boolean) => void } = { resolve: () => {} };
  const firstWaitPromise = new Promise<boolean>((resolve) => {
    firstWait.resolve = resolve;
  });
  let notifyFirstWaitStarted: () => void = () => {};
  const firstWaitStarted = new Promise<void>((resolve) => {
    notifyFirstWaitStarted = resolve;
  });
  let waits = 0;
  const harness = createHarness({
    windowReady: () => {
      waits += 1;
      if (waits === 1) {
        notifyFirstWaitStarted();
        return firstWaitPromise;
      }
      return Promise.resolve(true);
    },
  });

  harness.setMediaPath('https://127.0.0.1/one.m3u8');
  const first = harness.runtime.handleMediaPathChange('https://127.0.0.1/one.m3u8');
  await firstWaitStarted;

  harness.setMediaPath('https://127.0.0.1/two.m3u8');
  await harness.runtime.handleMediaPathChange('https://127.0.0.1/two.m3u8');
  firstWait.resolve(true);
  await first;

  assert.equal(harness.calls.filter((call) => call === 'open-jimaku').length, 1);
});

test('closing Jimaku or failing to open it releases an owned pause', async () => {
  const closed = createHarness();
  closed.setMediaPath('https://127.0.0.1/one.m3u8');
  await closed.runtime.handleMediaPathChange('https://127.0.0.1/one.m3u8');
  closed.runtime.handleJimakuModalClosed();
  assert.equal(closed.calls.at(-1), 'pause:false');

  const failed = createHarness({ opened: false });
  failed.setMediaPath('https://127.0.0.1/two.m3u8');
  await failed.runtime.handleMediaPathChange('https://127.0.0.1/two.m3u8');
  assert.equal(failed.calls.at(-1), 'pause:false');
});

test('leaving Anime Browser playback releases an owned pause', async () => {
  const harness = createHarness({
    animeMedia: (mediaPath) => mediaPath.startsWith('https://127.0.0.1/'),
  });
  harness.setMediaPath('https://127.0.0.1/stream.m3u8');
  await harness.runtime.handleMediaPathChange('https://127.0.0.1/stream.m3u8');

  harness.setMediaPath('/video.mkv');
  await harness.runtime.handleMediaPathChange('/video.mkv');

  assert.equal(harness.calls.at(-1), 'pause:false');
});

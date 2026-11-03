import assert from 'node:assert/strict';
import test from 'node:test';

import { createJimakuAutoOpenRuntime } from './jimaku-auto-open';

const videoTrack = { type: 'video', id: 1, selected: true };
const englishOnlyTrackList = [
  videoTrack,
  { type: 'sub', id: 1, lang: 'eng', title: 'English', external: false, selected: true },
];
const japaneseTrackList = [
  ...englishOnlyTrackList,
  { type: 'sub', id: 2, lang: 'jpn', title: 'Japanese', external: false },
];
const unlabeledSidecarTrackList = [
  ...englishOnlyTrackList,
  { type: 'sub', id: 2, title: 'srt', external: true },
];

function makeHarness(
  overrides: {
    enabled?: boolean;
    mediaPath?: string | null;
    trackList?: unknown;
    trackListError?: boolean;
    primaryLanguages?: string[];
  } = {},
) {
  const state = {
    enabled: overrides.enabled ?? true,
    mediaPath: overrides.mediaPath === undefined ? '/videos/episode-01.mkv' : overrides.mediaPath,
    trackList: 'trackList' in overrides ? overrides.trackList : englishOnlyTrackList,
  };
  const events: string[] = [];
  const scheduled: Array<() => void> = [];

  const runtime = createJimakuAutoOpenRuntime({
    isEnabled: () => state.enabled,
    getCurrentMediaPath: () => state.mediaPath,
    getMpvClient: () => ({
      connected: true,
      requestProperty: async (name: string) => {
        if (name !== 'track-list' || overrides.trackListError) {
          throw new Error(`Unexpected property: ${name}`);
        }
        return state.trackList;
      },
    }),
    getPrimarySubtitleLanguages: () => overrides.primaryLanguages ?? ['ja', 'jpn'],
    holdPlayback: () => events.push('hold'),
    openJimaku: () => events.push('open'),
    schedule: (callback) => {
      scheduled.push(callback);
      return scheduled.length as never;
    },
    clearScheduled: (timer) => {
      scheduled[(timer as unknown as number) - 1] = () => {};
    },
  });

  const runScheduled = async (): Promise<void> => {
    for (const callback of scheduled.splice(0)) {
      callback();
    }
    await new Promise((resolve) => setTimeout(resolve, 0));
  };

  return { runtime, state, events, scheduled, runScheduled };
}

test('jimaku auto-open pauses playback and opens the modal when local media has no Japanese track', async () => {
  const harness = makeHarness();

  harness.runtime.handleMediaPathChange('/videos/episode-01.mkv');
  await harness.runScheduled();

  assert.deepEqual(harness.events, ['hold', 'open']);
});

for (const c of [
  { name: 'a Japanese track exists', trackList: japaneseTrackList },
  { name: 'a single unlabeled sidecar subtitle exists', trackList: unlabeledSidecarTrackList },
]) {
  test(`jimaku auto-open stays closed when ${c.name}`, async () => {
    const harness = makeHarness({ trackList: c.trackList });

    harness.runtime.handleMediaPathChange('/videos/episode-01.mkv');
    await harness.runScheduled();

    assert.deepEqual(harness.events, []);
  });
}

test('jimaku auto-open uses the configured primary subtitle languages', async () => {
  const harness = makeHarness({ primaryLanguages: ['eng'] });

  harness.runtime.handleMediaPathChange('/videos/episode-01.mkv');
  await harness.runScheduled();

  assert.deepEqual(harness.events, []);
});

for (const c of [
  { name: 'the toggle is off', overrides: { enabled: false } },
  {
    name: 'media is remote',
    overrides: { mediaPath: 'https://example.com/episode-01.mkv' },
    path: 'https://example.com/episode-01.mkv',
  },
  { name: 'the track list cannot be read', overrides: { trackListError: true } },
  { name: 'mpv returns no track list', overrides: { trackList: null } },
  {
    name: 'the file is audio with cover art',
    overrides: {
      trackList: [
        { type: 'audio', id: 1, selected: true },
        { type: 'video', id: 1, albumart: true },
      ],
    },
  },
]) {
  test(`jimaku auto-open does nothing when ${c.name}`, async () => {
    const harness = makeHarness(c.overrides);

    harness.runtime.handleMediaPathChange(c.path ?? '/videos/episode-01.mkv');
    await harness.runScheduled();

    assert.deepEqual(harness.events, []);
  });
}

test('jimaku auto-open checks each media load once and skips media that changed before the check', async () => {
  const harness = makeHarness();

  harness.runtime.handleMediaPathChange('/videos/episode-01.mkv');
  harness.runtime.handleMediaPathChange('/videos/episode-01.mkv');
  harness.state.mediaPath = '/videos/episode-02.mkv';
  harness.runtime.handleMediaPathChange('/videos/episode-02.mkv');
  await harness.runScheduled();

  assert.deepEqual(harness.events, ['hold', 'open']);

  harness.runtime.handleMediaPathChange('/videos/episode-02.mkv');
  await harness.runScheduled();

  assert.deepEqual(harness.events, ['hold', 'open']);
});

test('jimaku auto-open holds playback until the modal closes or the media changes', async () => {
  const harness = makeHarness();
  assert.equal(harness.runtime.isHoldingPlayback(), false);

  harness.runtime.handleMediaPathChange('/videos/episode-01.mkv');
  await harness.runScheduled();
  assert.equal(harness.runtime.isHoldingPlayback(), true);

  harness.runtime.handleJimakuModalClosed();
  assert.equal(harness.runtime.isHoldingPlayback(), false);

  harness.state.mediaPath = '/videos/episode-02.mkv';
  harness.runtime.handleMediaPathChange('/videos/episode-02.mkv');
  await harness.runScheduled();
  assert.equal(harness.runtime.isHoldingPlayback(), true);

  harness.runtime.handleMediaPathChange('/videos/episode-03.mkv');
  assert.equal(harness.runtime.isHoldingPlayback(), false);
});

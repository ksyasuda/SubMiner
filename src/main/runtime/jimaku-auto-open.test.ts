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

const settle = () => new Promise((resolve) => setTimeout(resolve, 0));

function makeHarness(
  overrides: {
    enabled?: boolean;
    mediaPath?: string | null;
    trackList?: unknown;
    trackListError?: boolean;
    primaryLanguages?: string[];
    openJimaku?: () => Promise<boolean>;
  } = {},
) {
  const state = {
    enabled: overrides.enabled ?? true,
    mediaPath: overrides.mediaPath === undefined ? '/videos/episode-01.mkv' : overrides.mediaPath,
    trackList: 'trackList' in overrides ? overrides.trackList : englishOnlyTrackList,
    paused: false as boolean | null,
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
    isPlaybackPaused: () => state.paused,
    pausePlayback: () => {
      state.paused = true;
      events.push('pause');
    },
    endHold: ({ undoPause }) => {
      events.push('end-hold');
      if (undoPause) {
        state.paused = false;
        events.push('resume');
      }
    },
    holdPluginAutoplayGate: () => events.push('plugin-hold'),
    openJimaku: () => {
      events.push('open');
      return overrides.openJimaku?.() ?? Promise.resolve(true);
    },
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
    await settle();
  };
  // Reports a media load and runs the delayed check to completion.
  const load = async (mediaPath = state.mediaPath): Promise<void> => {
    state.mediaPath = mediaPath;
    runtime.handleMediaPathChange(mediaPath);
    await runScheduled();
  };
  const count = (event: string) => events.filter((entry) => entry === event).length;

  return { runtime, state, events, load, runScheduled, count };
}

test('jimaku auto-open pauses, then opens the modal, for local video with no Japanese track', async () => {
  const h = makeHarness();

  await h.load();

  assert.ok(h.events.indexOf('pause') < h.events.indexOf('open'));
  assert.equal(h.state.paused, true);
  assert.equal(h.count('plugin-hold'), 1);
  assert.equal(h.runtime.isHoldingPlayback(), true);
});

for (const c of [
  { name: 'a Japanese track exists', overrides: { trackList: japaneseTrackList } },
  {
    name: 'a single unlabeled sidecar subtitle exists',
    overrides: { trackList: unlabeledSidecarTrackList },
  },
  { name: 'the configured primary language matches', overrides: { primaryLanguages: ['eng'] } },
  { name: 'the toggle is off', overrides: { enabled: false } },
  { name: 'media is remote', overrides: { mediaPath: 'https://example.com/episode-01.mkv' } },
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
    const h = makeHarness(c.overrides);

    await h.load();

    assert.deepEqual(h.events, []);
    assert.equal(h.runtime.isHoldingPlayback(), false);
  });
}

test('jimaku auto-open checks each media load once and skips media that changed before the check', async () => {
  const h = makeHarness();

  h.runtime.handleMediaPathChange('/videos/episode-01.mkv');
  await h.load('/videos/episode-02.mkv');
  await h.load('/videos/episode-02.mkv');

  assert.equal(h.count('open'), 1);
});

test('closing the auto-opened modal ends the hold and undoes its pause', async () => {
  const h = makeHarness();
  await h.load();

  h.runtime.handleJimakuModalClosed();

  assert.equal(h.runtime.isHoldingPlayback(), false);
  assert.equal(h.count('end-hold'), 1);
  assert.equal(h.state.paused, false);
});

test('a pause that was already there is left to autoplay readiness when the modal closes', async () => {
  const h = makeHarness();
  h.state.paused = true; // e.g. the plugin's pause-until-ready gate, or the user

  await h.load();
  assert.equal(h.count('pause'), 0);
  assert.equal(h.runtime.isHoldingPlayback(), true);

  h.runtime.handleJimakuModalClosed();
  assert.equal(h.count('end-hold'), 1);
  assert.equal(h.state.paused, true);
});

test('changing media ends the hold without resuming the new media', async () => {
  const h = makeHarness();
  await h.load();

  h.state.trackList = japaneseTrackList;
  await h.load('/videos/episode-02.mkv');

  assert.equal(h.runtime.isHoldingPlayback(), false);
  assert.equal(h.count('resume'), 0);
});

for (const c of [
  { name: 'resolves false', openJimaku: () => Promise.resolve(false) },
  { name: 'rejects', openJimaku: () => Promise.reject(new Error('overlay gone')) },
]) {
  test(`when opening ${c.name}, jimaku auto-open undoes its own pause and leaves the plugin gate alone`, async () => {
    const h = makeHarness({ openJimaku: c.openJimaku });

    await h.load();

    assert.equal(h.runtime.isHoldingPlayback(), false);
    assert.equal(h.state.paused, false);
    assert.equal(h.count('plugin-hold'), 0);
  });
}

test('a failed open keeps a pause that was already there', async () => {
  const h = makeHarness({ openJimaku: () => Promise.resolve(false) });
  h.state.paused = true;

  await h.load();

  assert.equal(h.count('resume'), 0);
  assert.equal(h.state.paused, true);
});

test('a modal that acknowledges after the open timed out re-applies the hold', async () => {
  const h = makeHarness({ openJimaku: () => Promise.resolve(false) });
  await h.load();
  assert.equal(h.state.paused, false);

  h.runtime.handleJimakuModalOpened();

  assert.equal(h.runtime.isHoldingPlayback(), true);
  assert.equal(h.state.paused, true);
  assert.equal(h.count('plugin-hold'), 1);
});

test("a Jimaku modal opened after the late-ack window is the user's own and does not hold", async () => {
  const h = makeHarness({ openJimaku: () => Promise.resolve(false) });
  await h.load();
  await h.runScheduled(); // the late-ack window expires

  h.runtime.handleJimakuModalOpened();

  assert.equal(h.runtime.isHoldingPlayback(), false);
  assert.equal(h.count('plugin-hold'), 0);
});

test('a Jimaku modal opened without a pending auto-open does not hold playback', async () => {
  const h = makeHarness({ trackList: japaneseTrackList });
  await h.load();

  h.runtime.handleJimakuModalOpened();

  assert.equal(h.runtime.isHoldingPlayback(), false);
  assert.deepEqual(h.events, []);
});

test('jimaku auto-open ignores a late open result for media that already changed', async () => {
  const pendingOpens: Array<(opened: boolean) => void> = [];
  const h = makeHarness({
    openJimaku: () =>
      new Promise<boolean>((resolve) => {
        pendingOpens.push(resolve);
      }),
  });

  await h.load();
  await h.load('/videos/episode-02.mkv');
  const resumesBefore = h.count('resume');

  pendingOpens[0]?.(false);
  await settle();

  assert.equal(h.runtime.isHoldingPlayback(), true);
  assert.equal(h.count('resume'), resumesBefore);
});

test('a close reported before the auto-open is acknowledged belongs to the previous modal', async () => {
  const pendingOpens: Array<(opened: boolean) => void> = [];
  const h = makeHarness({
    openJimaku: () =>
      new Promise<boolean>((resolve) => {
        pendingOpens.push(resolve);
      }),
  });
  await h.load(); // open sent, not yet acknowledged

  h.runtime.handleJimakuModalClosed();
  assert.equal(h.runtime.isHoldingPlayback(), true);
  assert.equal(h.count('end-hold'), 0);

  h.runtime.handleJimakuModalOpened();
  pendingOpens[0]?.(true);
  await settle();
  h.runtime.handleJimakuModalClosed();
  assert.equal(h.runtime.isHoldingPlayback(), false);
  assert.equal(h.state.paused, false);
});

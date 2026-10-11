import assert from 'node:assert/strict';
import test from 'node:test';
import type { JellyfinPlaybackPlan } from '../../core/services/jellyfin';
import type { JellyfinTimelinePlaybackState } from '../../core/services/jellyfin-playback-reporter';
import { createPlayJellyfinItemInMpvHandler } from './jellyfin-playback-launch';

type Deps = Parameters<typeof createPlayJellyfinItemInMpvHandler>[0];
type PlayParams = Parameters<ReturnType<typeof createPlayJellyfinItemInMpvHandler>>[0];

const baseSession = {
  serverUrl: 'http://localhost:8096',
  accessToken: 'token',
  userId: 'uid',
  username: 'alice',
};

const baseClientInfo = {
  clientName: 'SubMiner',
  clientVersion: '1.0.0',
  deviceId: 'did',
};

const SUPPRESSED_SUBTITLE_OPTIONS =
  'sid=no,secondary-sid=no,sub-auto=no,sub-visibility=no,secondary-sub-visibility=no';

function makePlan(overrides: Partial<JellyfinPlaybackPlan> = {}): JellyfinPlaybackPlan {
  return {
    url: 'https://stream.example/video.m3u8',
    mode: 'direct',
    title: 'Episode 1',
    itemTitle: 'Episode 1',
    seriesTitle: null,
    seasonNumber: null,
    episodeNumber: null,
    startTimeTicks: 0,
    audioStreamIndex: null,
    subtitleStreamIndex: null,
    ...overrides,
  };
}

/**
 * Builds the playback handler with recording fakes. `events` is one ordered log
 * of everything the handler did, for the few tests where order is the behavior.
 */
function makeHarness(options: { plan?: Partial<JellyfinPlaybackPlan>; deps?: Partial<Deps> } = {}) {
  const commands: Array<Array<string | number>> = [];
  const events: string[] = [];
  const activeStates: Array<Parameters<Deps['setActivePlayback']>[0]> = [];
  const reports: JellyfinTimelinePlaybackState[] = [];
  const reporterParams: Array<Parameters<Deps['createPlaybackReporter']>[0]> = [];
  const reporter = {
    reportPlaying: async (state: JellyfinTimelinePlaybackState) => void reports.push(state),
    reportProgress: async () => {},
    reportStopped: async () => {},
  };

  const handler = createPlayJellyfinItemInMpvHandler({
    ensureMpvConnectedForPlayback: async () => true,
    getMpvClient: () => ({ connected: true, send: () => {} }),
    resolvePlaybackPlan: async () => makePlan(options.plan),
    applyJellyfinMpvDefaults: () => {},
    showVisibleOverlay: () => events.push('visible-overlay'),
    sendMpvCommand: (command) => {
      commands.push(command);
      events.push(`cmd:${command[0]}`);
    },
    armQuitOnDisconnect: () => events.push('arm'),
    schedule: () => {},
    convertTicksToSeconds: (ticks) => ticks / 10_000_000,
    preloadExternalSubtitles: () => {},
    setActivePlayback: (state) => {
      activeStates.push(state);
      events.push(`active:${String(state.loadedMediaPath)}`);
    },
    setLastProgressAtMs: () => {},
    createPlaybackReporter: (params) => {
      reporterParams.push(params);
      return reporter;
    },
    showMpvOsd: () => {},
    ...options.deps,
  });

  return {
    commands,
    events,
    activeStates,
    reports,
    reporter,
    reporterParams,
    play: (params: Partial<PlayParams> = {}) =>
      handler({
        session: baseSession,
        clientInfo: baseClientInfo,
        jellyfinConfig: {},
        itemId: 'item-1',
        ...params,
      }),
    loadfile: () => commands.find((command) => command[0] === 'loadfile'),
    loadedUrl: () => new URL(String(commands.find((command) => command[0] === 'loadfile')?.[1])),
  };
}

test('playback handler throws when mpv is not connected', async () => {
  const harness = makeHarness({
    deps: {
      ensureMpvConnectedForPlayback: async () => false,
      getMpvClient: () => null,
      resolvePlaybackPlan: async () => {
        throw new Error('unreachable');
      },
    },
  });

  await assert.rejects(() => harness.play(), /MPV not connected and auto-launch failed/);
});

test('playback handler disables mpv subtitle selection, then loads media and reports playback', async () => {
  const statsMetadata: unknown[] = [];
  const harness = makeHarness({
    plan: {
      title: 'Episode 1',
      itemTitle: 'Episode 1',
      seriesTitle: 'Show Title',
      seasonNumber: 1,
      episodeNumber: 1,
      startTimeTicks: 12_000_000,
      audioStreamIndex: 1,
      subtitleStreamIndex: 2,
    },
    deps: { recordJellyfinPlaybackMetadata: (metadata) => void statsMetadata.push(metadata) },
  });

  await harness.play();

  assert.deepEqual(harness.commands, [
    ['set_property', 'sub-auto', 'no'],
    ['set_property', 'sid', 'no'],
    ['set_property', 'secondary-sid', 'no'],
    ['set_property', 'sub-visibility', 'no'],
    ['set_property', 'secondary-sub-visibility', 'no'],
    ['script-message', 'subminer-managed-subtitles-loading'],
    ['set_property', 'force-media-title', 'Episode 1'],
    [
      'loadfile',
      'https://stream.example/video.m3u8',
      'replace',
      -1,
      `${SUPPRESSED_SUBTITLE_OPTIONS},start=1.2`,
    ],
  ]);
  assert.equal(harness.activeStates.length, 1);
  assert.equal(harness.activeStates[0]?.playMethod, 'DirectPlay');
  assert.equal(harness.activeStates[0]?.lastKnownPositionSeconds, 1.2);
  assert.equal(harness.activeStates[0]?.reporter, harness.reporter);
  assert.deepEqual(harness.reporterParams, [{ session: baseSession, clientInfo: baseClientInfo }]);
  assert.equal(harness.reports.length, 1);
  assert.equal(harness.reports[0]?.eventName, 'start');
  assert.equal(harness.reports[0]?.positionTicks, 12_000_000);
  assert.equal(harness.reports[0]?.isPaused, false);
  assert.deepEqual(statsMetadata, [
    {
      mediaPath: 'https://stream.example/video.m3u8',
      displayTitle: 'Episode 1',
      itemTitle: 'Episode 1',
      seriesTitle: 'Show Title',
      seasonNumber: 1,
      episodeNumber: 1,
      itemId: 'item-1',
    },
  ]);
  assert.ok(harness.events.includes('arm'));
});

test('playback handler waits for Jellyfin subtitle preload before showing visible overlay', async () => {
  const calls: string[] = [];
  let resolvePreload!: () => void;
  const preloadComplete = new Promise<void>((resolve) => {
    resolvePreload = resolve;
  });
  const harness = makeHarness({
    deps: {
      showVisibleOverlay: () => calls.push('visible-overlay'),
      preloadExternalSubtitles: async () => {
        calls.push('preload-start');
        await preloadComplete;
        calls.push('preload-done');
      },
    },
  });

  const playback = harness.play();
  for (let i = 0; i < 5 && calls.length === 0; i += 1) {
    await Promise.resolve();
  }

  assert.deepEqual(calls, ['preload-start']);
  resolvePreload();
  await playback;

  assert.deepEqual(calls, ['preload-start', 'preload-done', 'visible-overlay']);
});

test('playback handler strips Jellyfin subtitle stream from mpv load URL', async () => {
  const harness = makeHarness({
    plan: {
      url: 'https://jellyfin.local/Videos/ep-1/stream?static=true&api_key=secret-token&MediaSourceId=ms-1&AudioStreamIndex=3&SubtitleStreamIndex=4',
      audioStreamIndex: 3,
      subtitleStreamIndex: 4,
    },
  });

  await harness.play({ itemId: 'ep-1' });

  const url = harness.loadedUrl();
  assert.equal(url.searchParams.get('AudioStreamIndex'), '3');
  assert.equal(url.searchParams.has('SubtitleStreamIndex'), false);
  assert.equal(harness.reports[0]?.subtitleStreamIndex, 4);
});

const START_TICKS_CASES: Array<{
  name: string;
  planTicks: number;
  params: Partial<PlayParams>;
  expectedUrlTicks: string | null;
  expectedPositionTicks: number;
}> = [
  {
    name: 'keeps plan resume ticks when no override is given',
    planTicks: 35_000_000,
    params: {},
    expectedUrlTicks: '35000000',
    expectedPositionTicks: 35_000_000,
  },
  {
    name: 'applies a positive override to the stream url for remote resume',
    planTicks: 0,
    params: { startTimeTicksOverride: 55_000_000 },
    expectedUrlTicks: '55000000',
    expectedPositionTicks: 55_000_000,
  },
  {
    name: 'starts from the beginning on a zero override despite saved plan progress',
    planTicks: 35_000_000,
    params: { startTimeTicksOverride: 0, fallbackToPlanStartTimeOnZeroOverride: false },
    expectedUrlTicks: null,
    expectedPositionTicks: 0,
  },
  {
    name: 'keeps plan resume ticks on a zero override when fallback is enabled',
    planTicks: 35_000_000,
    params: { startTimeTicksOverride: 0, fallbackToPlanStartTimeOnZeroOverride: true },
    expectedUrlTicks: '35000000',
    expectedPositionTicks: 35_000_000,
  },
];

for (const c of START_TICKS_CASES) {
  test(`playback handler ${c.name}`, async () => {
    const harness = makeHarness({
      plan: {
        url: `https://stream.example/video.m3u8?api_key=token${
          c.planTicks > 0 ? `&StartTimeTicks=${c.planTicks}` : ''
        }`,
        mode: 'transcode',
        startTimeTicks: c.planTicks,
      },
    });

    await harness.play(c.params);

    assert.equal(harness.loadedUrl().searchParams.get('StartTimeTicks'), c.expectedUrlTicks);
    assert.equal(harness.reports[0]?.positionTicks, c.expectedPositionTicks);
    assert.equal(
      harness.commands.some((command) => command[0] === 'seek'),
      false,
    );
  });
}

test('playback handler publishes Jellyfin title before loading tokenized stream url', async () => {
  const harness = makeHarness({
    plan: {
      url: 'https://jellyfin.local/Videos/ep-1/stream?static=true&api_key=secret-token&MediaSourceId=ms-1',
      title: 'Galaxy Quest S02E07 A New Hope',
    },
    deps: {
      updateCurrentMediaTitle: (title) => void harness.events.push(`title:${title}`),
    },
  });

  await harness.play({ itemId: 'ep-1' });

  const titleIndex = harness.events.indexOf('title:Galaxy Quest S02E07 A New Hope');
  const loadIndex = harness.events.indexOf('cmd:loadfile');
  assert.ok(titleIndex >= 0);
  assert.ok(titleIndex < loadIndex);
  const mpvTitleIndex = harness.commands.findIndex((command) => command[1] === 'force-media-title');
  assert.ok(
    mpvTitleIndex >= 0 && mpvTitleIndex < harness.commands.findIndex((c) => c[0] === 'loadfile'),
  );
});

test('playback handler arms unloaded active playback before loading mpv media', async () => {
  const harness = makeHarness();

  await harness.play();

  const activeIndex = harness.events.indexOf('active:null');
  assert.ok(activeIndex >= 0);
  assert.ok(activeIndex < harness.events.indexOf('cmd:loadfile'));
});

const HOOK_FAILURE_CASES: Array<{ name: string; deps: Partial<Deps> }> = [
  {
    name: 'stats metadata failures',
    deps: {
      recordJellyfinPlaybackMetadata: () => {
        throw new Error('stats db unavailable');
      },
    },
  },
  {
    name: 'media title failures',
    deps: {
      updateCurrentMediaTitle: () => {
        throw new Error('title state unavailable');
      },
    },
  },
  {
    name: 'rejected best-effort hook promises',
    deps: {
      updateCurrentMediaTitle: async () => {
        throw new Error('title async unavailable');
      },
      recordJellyfinPlaybackMetadata: async () => {
        throw new Error('stats async unavailable');
      },
    },
  },
];

for (const c of HOOK_FAILURE_CASES) {
  test(`playback handler does not let ${c.name} block playback startup`, async () => {
    const harness = makeHarness({ deps: c.deps });

    await harness.play();
    await Promise.resolve();
    await Promise.resolve();

    assert.deepEqual(harness.loadfile(), [
      'loadfile',
      'https://stream.example/video.m3u8',
      'replace',
      -1,
      SUPPRESSED_SUBTITLE_OPTIONS,
    ]);
  });
}

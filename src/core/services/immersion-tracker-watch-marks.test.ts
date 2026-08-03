import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

type ImmersionTrackerService = import('./immersion-tracker-service').ImmersionTrackerService;

const POLICY = {
  batchSize: 10,
  flushIntervalMs: 5_000,
  queueCap: 100,
  payloadCapBytes: 512,
  maintenanceIntervalMs: 60 * 60 * 1000,
  retention: {
    eventsDays: 14,
    telemetryDays: 45,
    sessionsDays: 60,
    dailyRollupsDays: 730,
    monthlyRollupsDays: 3650,
    vacuumIntervalDays: 14,
  },
};

async function createTracker(): Promise<{ tracker: ImmersionTrackerService; dir: string }> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-watch-marks-'));
  const { ImmersionTrackerService: Ctor } = await import('./immersion-tracker-service');
  return { tracker: new Ctor({ dbPath: path.join(dir, 'immersion.sqlite'), policy: POLICY }), dir };
}

function episode(statsPath: string, episodeNumber: number) {
  return {
    mediaPath: '',
    statsPath,
    displayTitle: `Test Series S03E0${episodeNumber}`,
    seriesTitle: 'Test Series',
    seasonNumber: 3,
    episodeNumber,
  };
}

const EP1 = 'animebrowser://src/anime/ep1';
const EP2 = 'animebrowser://src/anime/ep2';

test('marking an episode nobody played records it, and clearing it takes the mark away', async () => {
  const { tracker, dir } = await createTracker();
  try {
    assert.equal(await tracker.setStreamWatchState([episode(EP1, 1), episode(EP2, 2)], true), 2);

    const marked = await tracker.getStreamWatchState([EP1, EP2]);
    assert.equal(marked.get(EP1)?.watched, true);
    assert.equal(marked.get(EP2)?.watched, true);
    // Nothing was played, so there is no session behind the mark.
    assert.equal(marked.get(EP1)?.sessionCount, 0);
    assert.equal(marked.get(EP1)?.lastWatchedMs, null);

    assert.equal(await tracker.setStreamWatchState([episode(EP1, 1)], false), 1);
    const cleared = await tracker.getStreamWatchState([EP1, EP2]);
    assert.equal(cleared.get(EP1)?.watched, false);
    assert.equal(cleared.get(EP2)?.watched, true);
  } finally {
    tracker.destroy();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('clearing a mark on an episode with no history creates nothing', async () => {
  const { tracker, dir } = await createTracker();
  try {
    assert.equal(await tracker.setStreamWatchState([episode(EP1, 1)], false), 0);
    assert.equal((await tracker.getStreamWatchState([EP1])).size, 0);
  } finally {
    tracker.destroy();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a manual mark stays out of the stats library until the episode is watched', async () => {
  const { tracker, dir } = await createTracker();
  try {
    await tracker.setStreamWatchState([episode(EP1, 1)], true);
    // Both library views join the lifetime tables, which only playback fills.
    assert.deepEqual(await tracker.getMediaLibrary(), []);
    assert.deepEqual(await tracker.getAnimeLibrary(), []);
  } finally {
    tracker.destroy();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a batch rebuilds the lifetime summaries once, not once per episode', async () => {
  const { tracker, dir } = await createTracker();
  const privateApi = tracker as unknown as {
    recordPrePlaybackMetadata: (params: { deferLifetimeRebuild?: boolean }) => boolean;
  };
  const original = privateApi.recordPrePlaybackMetadata.bind(tracker);
  let deferred = 0;
  let immediate = 0;
  privateApi.recordPrePlaybackMetadata = (params) => {
    if (params.deferLifetimeRebuild) deferred += 1;
    else immediate += 1;
    return original(params);
  };

  try {
    const season = Array.from({ length: 12 }, (_, index) =>
      episode(`animebrowser://src/anime/batch-${index}`, index + 1),
    );
    await tracker.setStreamWatchState(season, true);

    assert.equal(deferred, 12);
    assert.equal(immediate, 0, 'every row in the batch defers its rebuild to the end of the call');
  } finally {
    tracker.destroy();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('an episode with no stats path is skipped rather than recorded as unknown', async () => {
  const { tracker, dir } = await createTracker();
  try {
    assert.equal(await tracker.setStreamWatchState([episode('  ', 1)], true), 0);
  } finally {
    tracker.destroy();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test, { after } from 'node:test';
import { startFakeJellyfin, type FakeJellyfin } from './harness/fake-jellyfin';
import { generateFixtureClip } from './harness/fixtures';
import { startE2eSession, type E2eSession } from './harness/session';
import { waitUntil } from './harness/wait';

const TEST_TIMEOUT_MS = 180_000;
const EPISODE_IDS = ['ep1', 'ep2', 'ep3'];

let clipsDir: string | undefined;
let jellyfin: FakeJellyfin | undefined;
let session: E2eSession | undefined;

after(async () => {
  await session?.dispose();
  await jellyfin?.stop();
  if (clipsDir) fs.rmSync(clipsDir, { recursive: true, force: true });
});

type PlaylistEntry = { filename: string; current?: boolean; title?: string };

// The episode ids in mpv's playlist, the current one marked with `*`.
async function playlistIds(mpv: E2eSession['mpv']): Promise<string[]> {
  const playlist = (await mpv.command('get_property', 'playlist')) as PlaylistEntry[];
  return playlist.map((entry) => {
    const id = /\/Videos\/([^/]+)\/stream/.exec(entry.filename)?.[1] ?? entry.filename;
    return entry.current ? `${id}*` : id;
  });
}

// `playlist/N/title` is writable from mpv 0.42.
async function supportsPlaylistTitles(mpv: E2eSession['mpv']): Promise<boolean> {
  const version = String(await mpv.command('get_property', 'mpv-version'));
  const [major = 0, minor = 0] = (/v?(\d+)\.(\d+)/.exec(version) ?? []).slice(1).map(Number);
  return major > 0 || minor >= 42;
}

async function mainPid(started: E2eSession): Promise<number> {
  const main = await started.mainProcess();
  try {
    return await main.evaluate<number>('process.pid');
  } finally {
    main.close();
  }
}

test(
  'Jellyfin playback moves through queued series episodes in the same session',
  { timeout: TEST_TIMEOUT_MS },
  async () => {
    clipsDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-e2e-jellyfin-'));
    const episodes = EPISODE_IDS.map((id, index) => ({
      id,
      name: `Episode ${index + 1}`,
      mediaPath: path.join(clipsDir!, `${id}.mkv`),
    }));
    await Promise.all(episodes.map((episode) => generateFixtureClip(episode.mediaPath)));
    jellyfin = await startFakeJellyfin(episodes);
    session = await startE2eSession({
      config: {
        jellyfin: {
          enabled: true,
          serverUrl: jellyfin.url,
          remoteControlEnabled: false,
          episodePlaylist: { enabled: true, maxEpisodes: 3 },
        },
      },
      env: {
        SUBMINER_JELLYFIN_ACCESS_TOKEN: jellyfin.accessToken,
        SUBMINER_JELLYFIN_USER_ID: jellyfin.userId,
      },
    });
    const { mpv } = session;
    const pid = await mainPid(session);
    const reports = jellyfin.reports;
    const reported = (kind: string, itemId: string) =>
      reports.findIndex((report) => report.kind === kind && report.itemId === itemId);

    await session.runAppCommand('--jellyfin-play', '--jellyfin-item-id=ep2');
    await waitUntil(async () => (await playlistIds(mpv)).join() === 'ep1,ep2*,ep3', {
      description: 'the episodes around ep2 queued in the mpv playlist',
      timeoutMs: 30_000,
    });
    // The start report is fire-and-forget, so it can land after the playlist fills.
    await waitUntil(() => reported('playing', 'ep2') >= 0, {
      description: 'Jellyfin to hear ep2 start',
      timeoutMs: 30_000,
    });
    // mpv's playlist lists entry titles, not stream URLs carrying the API key. Entry titles
    // became writable in mpv 0.42; older players keep listing URLs.
    const titles = [
      'E2E Show S01E01 Episode 1',
      'E2E Show S01E02 Episode 2',
      'E2E Show S01E03 Episode 3',
    ];
    if (await supportsPlaylistTitles(mpv)) {
      await waitUntil(
        async () => {
          const playlist = (await mpv.command('get_property', 'playlist')) as PlaylistEntry[];
          return playlist.map((entry) => entry.title).join('|') === titles.join('|');
        },
        {
          description: 'queued playlist entries titled like the playing episode',
          timeoutMs: 10_000,
        },
      );
    }

    await mpv.command('playlist-next');
    await waitUntil(() => reported('playing', 'ep3') >= 0, {
      description: 'Jellyfin to hear ep3 start after playlist-next',
      timeoutMs: 30_000,
    });
    const stoppedEp2 = reported('stopped', 'ep2');
    assert.ok(stoppedEp2 >= 0, 'Jellyfin hears ep2 stop');
    assert.ok(stoppedEp2 < reported('playing', 'ep3'), 'ep2 stops before ep3 starts');
    assert.equal(await mpv.command('get_property', 'media-title'), 'E2E Show S01E03 Episode 3');
    // At the end of the series the window shifts back so it stays full.
    await waitUntil(async () => (await playlistIds(mpv)).join() === 'ep1,ep2,ep3*', {
      description: 'the playlist window re-centered on ep3',
      timeoutMs: 30_000,
    });

    await mpv.command('playlist-prev');
    await waitUntil(
      () =>
        reports.filter((report) => report.kind === 'playing' && report.itemId === 'ep2').length ===
        2,
      { description: 'Jellyfin to hear ep2 start again after playlist-prev', timeoutMs: 30_000 },
    );
    await waitUntil(async () => (await playlistIds(mpv)).join() === 'ep1,ep2*,ep3', {
      description: 'the playlist window re-centered on ep2',
      timeoutMs: 30_000,
    });

    assert.equal(await mainPid(session), pid, 'the app keeps running across episode switches');
  },
);

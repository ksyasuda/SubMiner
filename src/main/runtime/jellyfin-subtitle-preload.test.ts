import assert from 'node:assert/strict';
import test from 'node:test';
import { createPreloadJellyfinExternalSubtitlesHandler } from './jellyfin-subtitle-preload';

const session = {
  serverUrl: 'http://localhost:8096',
  accessToken: 'token',
  userId: 'uid',
  username: 'alice',
};

const clientInfo = {
  clientName: 'SubMiner',
  clientVersion: '1.0',
  deviceId: 'dev',
};

function makeDeps(overrides: {
  listJellyfinSubtitleTracks?: Parameters<
    typeof createPreloadJellyfinExternalSubtitlesHandler
  >[0]['listJellyfinSubtitleTracks'];
  getMpvClient?: Parameters<
    typeof createPreloadJellyfinExternalSubtitlesHandler
  >[0]['getMpvClient'];
  sendMpvCommand?: Parameters<
    typeof createPreloadJellyfinExternalSubtitlesHandler
  >[0]['sendMpvCommand'];
  wait?: Parameters<typeof createPreloadJellyfinExternalSubtitlesHandler>[0]['wait'];
  cacheSubtitleTrack?: Parameters<
    typeof createPreloadJellyfinExternalSubtitlesHandler
  >[0]['cacheSubtitleTrack'];
  cleanupCachedSubtitles?: Parameters<
    typeof createPreloadJellyfinExternalSubtitlesHandler
  >[0]['cleanupCachedSubtitles'];
  logDebug?: Parameters<typeof createPreloadJellyfinExternalSubtitlesHandler>[0]['logDebug'];
}) {
  return {
    listJellyfinSubtitleTracks: overrides.listJellyfinSubtitleTracks ?? (async () => []),
    getMpvClient: overrides.getMpvClient ?? (() => null),
    sendMpvCommand: overrides.sendMpvCommand ?? (() => {}),
    wait: overrides.wait ?? (async () => {}),
    cacheSubtitleTrack:
      overrides.cacheSubtitleTrack ??
      (async (track) => ({
        path: `/tmp/subminer-jellyfin-subtitles/${track.index}.srt`,
        cleanupDir: '/tmp/subminer-jellyfin-subtitles',
      })),
    cleanupCachedSubtitles: overrides.cleanupCachedSubtitles ?? (() => {}),
    logDebug: overrides.logDebug ?? (() => {}),
  };
}

test('preload jellyfin subtitles caches external tracks locally and chooses japanese+english tracks', async () => {
  const commands: Array<Array<string | number>> = [];
  const preload = createPreloadJellyfinExternalSubtitlesHandler(
    makeDeps({
      listJellyfinSubtitleTracks: async () => [
        { index: 0, language: 'jpn', title: 'Japanese', deliveryUrl: 'https://sub/a.srt' },
        { index: 1, language: 'eng', title: 'English SDH', deliveryUrl: 'https://sub/b.srt' },
        { index: 2, language: 'eng', title: 'English SDH', deliveryUrl: 'https://sub/b.srt' },
      ],
      getMpvClient: () => ({
        requestProperty: async () => [
          { type: 'sub', id: 5, lang: 'jpn', title: 'Japanese', external: true },
          { type: 'sub', id: 6, lang: 'eng', title: 'English', external: true },
        ],
      }),
      sendMpvCommand: (command) => commands.push(command),
      cacheSubtitleTrack: async (track) => ({
        path: `/tmp/subminer-jellyfin-subtitles/${track.index}.srt`,
        cleanupDir: '/tmp/subminer-jellyfin-subtitles',
      }),
    }),
  );

  await preload({ session, clientInfo, itemId: 'item-1' });

  assert.deepEqual(commands, [
    ['sub-add', '/tmp/subminer-jellyfin-subtitles/0.srt', 'cached', 'Japanese', 'jpn'],
    ['sub-add', '/tmp/subminer-jellyfin-subtitles/1.srt', 'cached', 'English SDH', 'eng'],
    ['set_property', 'sid', 5],
    ['set_property', 'secondary-sid', 6],
  ]);
});

test('preload jellyfin subtitles cleans previous cached subtitles before a new preload', async () => {
  const cleanupCalls: string[][] = [];
  const preload = createPreloadJellyfinExternalSubtitlesHandler(
    makeDeps({
      listJellyfinSubtitleTracks: async () => [
        { index: 0, language: 'jpn', title: 'Japanese', deliveryUrl: 'https://sub/a.srt' },
      ],
      getMpvClient: () => ({ requestProperty: async () => [] }),
      cacheSubtitleTrack: async (track) => ({
        path: `/tmp/subminer-jellyfin-subtitles-${track.index}/track.srt`,
        cleanupDir: `/tmp/subminer-jellyfin-subtitles-${track.index}`,
      }),
      cleanupCachedSubtitles: (dirs) => cleanupCalls.push(dirs),
    }),
  );

  await preload({ session, clientInfo, itemId: 'item-1' });
  await preload({ session, clientInfo, itemId: 'item-2' });

  assert.deepEqual(cleanupCalls, [['/tmp/subminer-jellyfin-subtitles-0']]);
});

test('preload jellyfin subtitles exposes cleanup for active cached subtitles', async () => {
  const cleanupCalls: string[][] = [];
  const preload = createPreloadJellyfinExternalSubtitlesHandler(
    makeDeps({
      listJellyfinSubtitleTracks: async () => [
        { index: 0, language: 'jpn', title: 'Japanese', deliveryUrl: 'https://sub/a.srt' },
      ],
      getMpvClient: () => ({ requestProperty: async () => [] }),
      cacheSubtitleTrack: async () => ({
        path: '/tmp/subminer-jellyfin-subtitles-active/track.srt',
        cleanupDir: '/tmp/subminer-jellyfin-subtitles-active',
      }),
      cleanupCachedSubtitles: (dirs) => cleanupCalls.push(dirs),
    }),
  );

  await preload({ session, clientInfo, itemId: 'item-1' });
  preload.cleanupCachedSubtitles();
  preload.cleanupCachedSubtitles();

  assert.deepEqual(cleanupCalls, [['/tmp/subminer-jellyfin-subtitles-active']]);
});

test('preload jellyfin subtitles exits quietly when no external tracks', async () => {
  const commands: Array<Array<string | number>> = [];
  let waited = false;
  const preload = createPreloadJellyfinExternalSubtitlesHandler(
    makeDeps({
      listJellyfinSubtitleTracks: async () => [{ index: 0, language: 'jpn', title: 'Embedded' }],
      getMpvClient: () => ({ requestProperty: async () => [] }),
      sendMpvCommand: (command) => commands.push(command),
      wait: async () => {
        waited = true;
      },
    }),
  );

  await preload({ session, clientInfo, itemId: 'item-1' });

  assert.equal(waited, false);
  assert.deepEqual(commands, []);
});

test('preload jellyfin subtitles logs debug on failure', async () => {
  const logs: string[] = [];
  const preload = createPreloadJellyfinExternalSubtitlesHandler(
    makeDeps({
      listJellyfinSubtitleTracks: async () => {
        throw new Error('network down');
      },
      getMpvClient: () => null,
      sendMpvCommand: () => {},
      wait: async () => {},
      logDebug: (message) => logs.push(message),
    }),
  );

  await preload({ session, clientInfo, itemId: 'item-1' });

  assert.deepEqual(logs, ['Failed to preload Jellyfin external subtitles']);
});

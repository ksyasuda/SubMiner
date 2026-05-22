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
          {
            type: 'sub',
            id: 5,
            lang: 'jpn',
            title: 'Japanese',
            external: true,
            'external-filename': '/tmp/subminer-jellyfin-subtitles/0.srt',
          },
          {
            type: 'sub',
            id: 6,
            lang: 'eng',
            title: 'English',
            external: true,
            'external-filename': '/tmp/subminer-jellyfin-subtitles/1.srt',
          },
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
    ['sub-add', '/tmp/subminer-jellyfin-subtitles/0.srt', 'auto', 'Japanese', 'jpn'],
    ['sub-add', '/tmp/subminer-jellyfin-subtitles/1.srt', 'auto', 'English SDH', 'eng'],
    ['set_property', 'sid', 5],
    ['set_property', 'secondary-sid', 6],
  ]);
});

test('preload jellyfin subtitles waits for delayed cached japanese track before selecting', async () => {
  const commands: Array<Array<string | number>> = [];
  let requestCount = 0;
  const preload = createPreloadJellyfinExternalSubtitlesHandler(
    makeDeps({
      listJellyfinSubtitleTracks: async () => [
        { index: 0, language: 'jpn', title: 'Japanese', deliveryUrl: 'https://sub/a.srt' },
        { index: 1, language: 'eng', title: 'English', deliveryUrl: 'https://sub/b.srt' },
      ],
      getMpvClient: () => ({
        requestProperty: async () => {
          requestCount += 1;
          if (requestCount < 3) {
            return [{ type: 'sub', id: 1, lang: 'eng', title: 'CR', external: false }];
          }
          return [
            { type: 'sub', id: 1, lang: 'eng', title: 'CR', external: false },
            {
              type: 'sub',
              id: 5,
              lang: 'jpn',
              title: 'Japanese',
              external: true,
              'external-filename': '/tmp/subminer-jellyfin-subtitles/0.srt',
            },
            {
              type: 'sub',
              id: 6,
              lang: 'eng',
              title: 'English',
              external: true,
              'external-filename': '/tmp/subminer-jellyfin-subtitles/1.srt',
            },
          ];
        },
      }),
      sendMpvCommand: (command) => commands.push(command),
    }),
  );

  await preload({ session, clientInfo, itemId: 'item-1' });

  assert.equal(requestCount, 3);
  assert.deepEqual(
    commands.filter((command) => command[0] === 'set_property'),
    [
      ['set_property', 'sid', 5],
      ['set_property', 'secondary-sid', 6],
    ],
  );
});

test('preload jellyfin subtitles waits for delayed external japanese track instead of embedded japanese', async () => {
  const commands: Array<Array<string | number>> = [];
  let requestCount = 0;
  const preload = createPreloadJellyfinExternalSubtitlesHandler(
    makeDeps({
      listJellyfinSubtitleTracks: async () => [
        { index: 0, language: 'jpn', title: 'Japanese', deliveryUrl: 'https://sub/a.srt' },
        { index: 1, language: 'eng', title: 'English', deliveryUrl: 'https://sub/b.srt' },
      ],
      getMpvClient: () => ({
        requestProperty: async () => {
          requestCount += 1;
          if (requestCount < 3) {
            return [{ type: 'sub', id: 2, lang: 'jpn', title: 'Embedded Japanese' }];
          }
          return [
            { type: 'sub', id: 2, lang: 'jpn', title: 'Embedded Japanese' },
            {
              type: 'sub',
              id: 42,
              lang: 'jpn',
              title: 'Japanese',
              external: true,
              'external-filename': '/tmp/subminer-jellyfin-subtitles/0.srt',
            },
            {
              type: 'sub',
              id: 43,
              lang: 'eng',
              title: 'English',
              external: true,
              'external-filename': '/tmp/subminer-jellyfin-subtitles/1.srt',
            },
          ];
        },
      }),
      sendMpvCommand: (command) => commands.push(command),
    }),
  );

  await preload({ session, clientInfo, itemId: 'item-1' });

  assert.equal(requestCount, 3);
  assert.deepEqual(
    commands.filter((command) => command[0] === 'set_property'),
    [
      ['set_property', 'sid', 42],
      ['set_property', 'secondary-sid', 43],
    ],
  );
});

test('preload jellyfin subtitles does not let later subtitle adds steal japanese primary selection', async () => {
  const commands: Array<Array<string | number>> = [];
  let requestCount = 0;
  const preload = createPreloadJellyfinExternalSubtitlesHandler(
    makeDeps({
      listJellyfinSubtitleTracks: async () => [
        { index: 1, language: 'jpn', title: 'Japanese', deliveryUrl: 'https://sub/jpn.srt' },
        { index: 10, language: 'deu', title: 'German', deliveryUrl: 'https://sub/deu.ass' },
        { index: 12, language: 'rus', title: 'Russian', deliveryUrl: 'https://sub/rus.ass' },
      ],
      getMpvClient: () => ({
        requestProperty: async () => {
          requestCount += 1;
          if (requestCount === 1) {
            return [
              {
                type: 'sub',
                id: 11,
                lang: 'jpn',
                title: 'Japanese',
                external: true,
                'external-filename': '/tmp/subminer-jellyfin-subtitles/1.srt',
              },
            ];
          }
          return [
            {
              type: 'sub',
              id: 11,
              lang: 'jpn',
              title: 'Japanese',
              external: true,
              'external-filename': '/tmp/subminer-jellyfin-subtitles/1.srt',
            },
            {
              type: 'sub',
              id: 18,
              lang: 'deu',
              title: 'German',
              external: true,
              selected: true,
              'external-filename': '/tmp/subminer-jellyfin-subtitles/10.srt',
            },
            {
              type: 'sub',
              id: 20,
              lang: 'rus',
              title: 'Russian',
              external: true,
              selected: true,
              'external-filename': '/tmp/subminer-jellyfin-subtitles/12.srt',
            },
          ];
        },
      }),
      sendMpvCommand: (command) => commands.push(command),
    }),
  );

  await preload({ session, clientInfo, itemId: 'item-1' });

  assert.equal(requestCount, 2);
  assert.deepEqual(
    commands.filter((command) => command[0] === 'sub-add'),
    [
      ['sub-add', '/tmp/subminer-jellyfin-subtitles/1.srt', 'auto', 'Japanese', 'jpn'],
      ['sub-add', '/tmp/subminer-jellyfin-subtitles/10.srt', 'auto', 'German', 'deu'],
      ['sub-add', '/tmp/subminer-jellyfin-subtitles/12.srt', 'auto', 'Russian', 'rus'],
    ],
  );
  assert.deepEqual(
    commands.filter((command) => command[0] === 'set_property'),
    [['set_property', 'sid', 11]],
  );
});

test('preload jellyfin subtitles leaves current track alone when reported japanese track never appears', async () => {
  const commands: Array<Array<string | number>> = [];
  const logs: string[] = [];
  let requestCount = 0;
  const preload = createPreloadJellyfinExternalSubtitlesHandler(
    makeDeps({
      listJellyfinSubtitleTracks: async () => [
        { index: 1, language: 'jpn', title: 'Japanese', deliveryUrl: 'https://sub/jpn.srt' },
      ],
      getMpvClient: () => ({
        requestProperty: async () => {
          requestCount += 1;
          return [{ type: 'sub', id: 1, lang: 'eng', title: 'CR', external: false }];
        },
      }),
      sendMpvCommand: (command) => commands.push(command),
      logDebug: (message) => logs.push(message),
    }),
  );

  await preload({ session, clientInfo, itemId: 'item-1' });

  assert.equal(requestCount, 10);
  assert.equal(
    commands.some(
      (command) => command[0] === 'set_property' && command[1] === 'sid' && command[2] === 'no',
    ),
    false,
  );
  assert.deepEqual(logs, ['Timed out waiting for Jellyfin Japanese subtitle track']);
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

test('preload jellyfin subtitles continues after cleanup failures', async () => {
  const commands: Array<Array<string | number>> = [];
  const logs: string[] = [];
  let cleanupShouldFail = false;
  const preload = createPreloadJellyfinExternalSubtitlesHandler(
    makeDeps({
      listJellyfinSubtitleTracks: async () => [
        { index: 0, language: 'eng', title: 'English', deliveryUrl: 'https://sub/a.srt' },
      ],
      getMpvClient: () => ({ requestProperty: async () => [] }),
      cacheSubtitleTrack: async (track) => ({
        path: `/tmp/subminer-jellyfin-subtitles-${track.index}/track.srt`,
        cleanupDir: `/tmp/subminer-jellyfin-subtitles-${track.index}`,
      }),
      sendMpvCommand: (command) => commands.push(command),
      cleanupCachedSubtitles: () => {
        if (cleanupShouldFail) {
          throw new Error('cleanup failed');
        }
      },
      logDebug: (message) => logs.push(message),
    }),
  );

  await preload({ session, clientInfo, itemId: 'item-1' });
  cleanupShouldFail = true;
  await assert.doesNotReject(() => preload({ session, clientInfo, itemId: 'item-2' }));

  assert.deepEqual(logs, ['Failed to cleanup Jellyfin cached subtitles']);
  assert.deepEqual(
    commands.filter((command) => command[0] === 'sub-add'),
    [
      ['sub-add', '/tmp/subminer-jellyfin-subtitles-0/track.srt', 'auto', 'English', 'eng'],
      ['sub-add', '/tmp/subminer-jellyfin-subtitles-0/track.srt', 'auto', 'English', 'eng'],
    ],
  );
});

test('preload jellyfin subtitles serializes overlapping preload runs', async () => {
  let releaseFirstList!: () => void;
  const firstListBlocked = new Promise<void>((resolve) => {
    releaseFirstList = resolve;
  });
  const listCalls: string[] = [];
  const preload = createPreloadJellyfinExternalSubtitlesHandler(
    makeDeps({
      listJellyfinSubtitleTracks: async (_session, _clientInfo, itemId) => {
        listCalls.push(itemId);
        if (itemId === 'item-1') {
          await firstListBlocked;
        }
        return [];
      },
    }),
  );

  const first = preload({ session, clientInfo, itemId: 'item-1' });
  const second = preload({ session, clientInfo, itemId: 'item-2' });
  await Promise.resolve();

  assert.deepEqual(listCalls, ['item-1']);
  releaseFirstList();
  await Promise.all([first, second]);
  assert.deepEqual(listCalls, ['item-1', 'item-2']);
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

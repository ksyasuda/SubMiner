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
  getSavedSubtitleDelay?: Parameters<
    typeof createPreloadJellyfinExternalSubtitlesHandler
  >[0]['getSavedSubtitleDelay'];
  setActiveSubtitleDelayKey?: Parameters<
    typeof createPreloadJellyfinExternalSubtitlesHandler
  >[0]['setActiveSubtitleDelayKey'];
  loadSubtitleSourceText?: (source: string) => Promise<string>;
  saveSubtitleDelay?: (itemId: string, streamIndex: number, delaySeconds: number) => void;
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
    getSavedSubtitleDelay: overrides.getSavedSubtitleDelay,
    setActiveSubtitleDelayKey: overrides.setActiveSubtitleDelayKey,
    loadSubtitleSourceText: overrides.loadSubtitleSourceText,
    saveSubtitleDelay: overrides.saveSubtitleDelay,
    logDebug: overrides.logDebug ?? (() => {}),
  };
}

function withoutTrackAutoSelectionCommands(
  commands: Array<Array<string | number>>,
): Array<Array<string | number>> {
  return commands.filter(
    (command) =>
      !(
        command[0] === 'set_property' &&
        (command[1] === 'track-auto-selection' ||
          (command[1] === 'sid' && command[2] === 'no') ||
          (command[1] === 'secondary-sid' && command[2] === 'no') ||
          (command[1] === 'sub-visibility' && command[2] === 'no') ||
          (command[1] === 'secondary-sub-visibility' && command[2] === 'no'))
      ),
  );
}

function setPropertyCommandsExceptTrackAutoSelection(
  commands: Array<Array<string | number>>,
): Array<Array<string | number>> {
  return withoutTrackAutoSelectionCommands(commands).filter(
    (command) => command[0] === 'set_property',
  );
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

  assert.deepEqual(withoutTrackAutoSelectionCommands(commands), [
    ['sub-add', '/tmp/subminer-jellyfin-subtitles/0.srt', 'auto', 'Japanese', 'jpn'],
    ['sub-add', '/tmp/subminer-jellyfin-subtitles/1.srt', 'auto', 'English SDH', 'eng'],
    ['set_property', 'sid', 5],
    ['set_property', 'secondary-sid', 6],
  ]);
});

test('preload jellyfin subtitles stages tracks without temporary subtitle selection', async () => {
  const commands: Array<Array<string | number>> = [];
  const preload = createPreloadJellyfinExternalSubtitlesHandler(
    makeDeps({
      listJellyfinSubtitleTracks: async () => [
        { index: 0, language: 'jpn', title: 'Japanese', deliveryUrl: 'https://sub/a.srt' },
        { index: 1, language: 'eng', title: 'English', deliveryUrl: 'https://sub/b.srt' },
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
    }),
  );

  await preload({ session, clientInfo, itemId: 'item-1' });

  assert.deepEqual(
    commands.filter((command) => command[0] === 'sub-add').map((command) => command[2]),
    ['auto', 'auto'],
  );
  const firstFinalSelectionIndex = commands.findIndex(
    (command) => command[0] === 'set_property' && command[1] === 'sid' && command[2] === 5,
  );
  assert.ok(firstFinalSelectionIndex >= 0);
  assert.equal(
    commands
      .slice(0, firstFinalSelectionIndex)
      .some(
        (command) =>
          command[0] === 'sub-add' && (command[2] === 'cached' || command[2] === 'select'),
      ),
    false,
  );
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
  assert.deepEqual(setPropertyCommandsExceptTrackAutoSelection(commands), [
    ['set_property', 'sid', 5],
    ['set_property', 'secondary-sid', 6],
  ]);
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
  assert.deepEqual(setPropertyCommandsExceptTrackAutoSelection(commands), [
    ['set_property', 'sid', 42],
    ['set_property', 'secondary-sid', 43],
  ]);
});

test('preload jellyfin subtitles prefers Jellyfin default and embedded japanese sources', async () => {
  const commands: Array<Array<string | number>> = [];
  const preload = createPreloadJellyfinExternalSubtitlesHandler(
    makeDeps({
      listJellyfinSubtitleTracks: async () => [
        {
          index: 0,
          language: 'jpn',
          title: 'External Japanese',
          isExternal: true,
          deliveryUrl: 'https://sub/external.srt',
        },
        {
          index: 1,
          language: 'jpn',
          title: 'Embedded Japanese',
          isDefault: true,
          isExternal: false,
          deliveryUrl: 'https://sub/embedded.srt',
        },
        {
          index: 2,
          language: 'eng',
          title: 'English',
          deliveryUrl: 'https://sub/english.srt',
        },
      ],
      getMpvClient: () => ({
        requestProperty: async () => [
          {
            type: 'sub',
            id: 5,
            lang: 'jpn',
            title: 'External Japanese',
            external: true,
            'external-filename': '/tmp/subminer-jellyfin-subtitles/0.srt',
          },
          {
            type: 'sub',
            id: 6,
            lang: 'jpn',
            title: 'Embedded Japanese',
            external: true,
            'external-filename': '/tmp/subminer-jellyfin-subtitles/1.srt',
          },
          {
            type: 'sub',
            id: 7,
            lang: 'eng',
            title: 'English',
            external: true,
            'external-filename': '/tmp/subminer-jellyfin-subtitles/2.srt',
          },
        ],
      }),
      sendMpvCommand: (command) => commands.push(command),
    }),
  );

  await preload({ session, clientInfo, itemId: 'item-1' });

  assert.deepEqual(setPropertyCommandsExceptTrackAutoSelection(commands), [
    ['set_property', 'sid', 6],
    ['set_property', 'secondary-sid', 7],
  ]);
});

test('preload jellyfin subtitles applies saved delay for selected japanese stream', async () => {
  const commands: Array<Array<string | number>> = [];
  const activeKeys: Array<{ itemId: string; streamIndex: number } | null> = [];
  const preload = createPreloadJellyfinExternalSubtitlesHandler(
    makeDeps({
      listJellyfinSubtitleTracks: async () => [
        { index: 3, language: 'jpn', title: 'Japanese', deliveryUrl: 'https://sub/jpn.srt' },
      ],
      getMpvClient: () => ({
        requestProperty: async () => [
          {
            type: 'sub',
            id: 11,
            lang: 'jpn',
            title: 'Japanese',
            external: true,
            'external-filename': '/tmp/subminer-jellyfin-subtitles/3.srt',
          },
        ],
      }),
      sendMpvCommand: (command) => commands.push(command),
      getSavedSubtitleDelay: (_itemId, streamIndex) => (streamIndex === 3 ? 1.25 : null),
      setActiveSubtitleDelayKey: (key) => activeKeys.push(key),
    }),
  );

  await preload({ session, clientInfo, itemId: 'item-9' });

  assert.deepEqual(setPropertyCommandsExceptTrackAutoSelection(commands), [
    ['set_property', 'sub-delay', 1.25],
    ['set_property', 'sid', 11],
  ]);
  assert.deepEqual(activeKeys, [{ itemId: 'item-9', streamIndex: 3 }]);
});

test('preload jellyfin subtitles applies saved delay before selecting japanese stream', async () => {
  const commands: Array<Array<string | number>> = [];
  const preload = createPreloadJellyfinExternalSubtitlesHandler(
    makeDeps({
      listJellyfinSubtitleTracks: async () => [
        { index: 3, language: 'jpn', title: 'Japanese', deliveryUrl: 'https://sub/jpn.srt' },
      ],
      getMpvClient: () => ({
        requestProperty: async () => [
          {
            type: 'sub',
            id: 11,
            lang: 'jpn',
            title: 'Japanese',
            external: true,
            'external-filename': '/tmp/subminer-jellyfin-subtitles/3.srt',
          },
        ],
      }),
      sendMpvCommand: (command) => commands.push(command),
      getSavedSubtitleDelay: () => 1.25,
    }),
  );

  await preload({ session, clientInfo, itemId: 'item-9' });

  const delayIndex = commands.findIndex(
    (command) => command[0] === 'set_property' && command[1] === 'sub-delay' && command[2] === 1.25,
  );
  const selectedSidIndex = commands.findIndex(
    (command) => command[0] === 'set_property' && command[1] === 'sid' && command[2] === 11,
  );
  assert.ok(delayIndex >= 0);
  assert.ok(selectedSidIndex >= 0);
  assert.ok(delayIndex < selectedSidIndex);
});

test('preload jellyfin subtitles auto-aligns late japanese track from english reference', async () => {
  const commands: Array<Array<string | number>> = [];
  const savedDelays: Array<{ itemId: string; streamIndex: number; delaySeconds: number }> = [];
  const primarySrt = `1
00:00:34,935 --> 00:00:36,937
Japanese 1

2
00:00:36,937 --> 00:00:41,441
Japanese 2

3
00:00:41,441 --> 00:00:45,279
Japanese 3

4
00:00:45,279 --> 00:00:48,115
Japanese 4

5
00:00:48,115 --> 00:00:52,286
Japanese 5

6
00:00:52,286 --> 00:00:54,955
Japanese 6

7
00:00:54,955 --> 00:00:59,793
Japanese 7

8
00:00:59,793 --> 00:01:03,630
Japanese 8

9
00:01:03,630 --> 00:01:07,634
Japanese 9

10
00:01:07,634 --> 00:01:13,040
Japanese 10

11
00:01:16,643 --> 00:01:20,814
Japanese 11

12
00:01:20,814 --> 00:01:23,116
Japanese 12

13
00:01:27,988 --> 00:01:30,991
Japanese 13

14
00:01:30,991 --> 00:01:34,094
Japanese 14

15
00:01:34,094 --> 00:01:37,097
Japanese 15

16
00:01:37,097 --> 00:01:39,100
Japanese 16
`;
  const referenceAss = `[Events]
Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
Dialogue: 0,0:00:03.46,0:00:08.73,Default,,0,0,0,,English 1
Dialogue: 0,0:00:09.48,0:00:13.61,Default,,0,0,0,,English 2
Dialogue: 0,0:00:13.61,0:00:19.64,Default,,0,0,0,,English 3
Dialogue: 0,0:00:21.40,0:00:27.32,Default,,0,0,0,,English 4
Dialogue: 0,0:00:28.16,0:00:31.75,Default,,0,0,0,,English 5
Dialogue: 0,0:00:32.06,0:00:34.52,Default,,0,0,0,,English 6
Dialogue: 0,0:00:35.93,0:00:40.57,Default,,0,0,0,,English 7
Dialogue: 0,0:00:45.10,0:00:51.01,Default,,0,0,0,,English 8
Dialogue: 0,0:00:56.57,0:00:59.12,Default,,0,0,0,,English 9
Dialogue: 0,0:00:59.68,0:01:02.44,Default,,0,0,0,,English 10
Dialogue: 0,0:01:02.44,0:01:05.56,Default,,0,0,0,,English 11
Dialogue: 0,0:01:05.56,0:01:06.87,Default,,0,0,0,,English 12
`;
  const preload = createPreloadJellyfinExternalSubtitlesHandler(
    makeDeps({
      listJellyfinSubtitleTracks: async () => [
        { index: 0, language: 'jpn', title: 'Japanese', deliveryUrl: 'https://sub/jpn.srt' },
        { index: 4, language: 'eng', title: 'English', deliveryUrl: 'https://sub/eng.ass' },
      ],
      getMpvClient: () => ({
        requestProperty: async () => [
          {
            type: 'sub',
            id: 10,
            lang: 'jpn',
            title: 'Japanese',
            external: true,
            'external-filename': '/tmp/subminer-jellyfin-subtitles/0.srt',
          },
          {
            type: 'sub',
            id: 12,
            lang: 'eng',
            title: 'English',
            external: true,
            'external-filename': '/tmp/subminer-jellyfin-subtitles/4.ass',
          },
        ],
      }),
      sendMpvCommand: (command) => commands.push(command),
      cacheSubtitleTrack: async (track) => ({
        path: `/tmp/subminer-jellyfin-subtitles/${track.index}.${track.index === 4 ? 'ass' : 'srt'}`,
        cleanupDir: '/tmp/subminer-jellyfin-subtitles',
      }),
      getSavedSubtitleDelay: () => null,
      loadSubtitleSourceText: async (source) =>
        source.endsWith('.ass') ? referenceAss : primarySrt,
      saveSubtitleDelay: (itemId, streamIndex, delaySeconds) => {
        savedDelays.push({ itemId, streamIndex, delaySeconds });
      },
    }),
  );

  await preload({ session, clientInfo, itemId: 'item-9' });

  const delayCommand = commands.find(
    (command) => command[0] === 'set_property' && command[1] === 'sub-delay',
  );
  assert.ok(delayCommand);
  const delaySeconds = delayCommand[2];
  if (typeof delaySeconds !== 'number') {
    assert.fail('Expected numeric subtitle delay.');
  }
  assert.ok(delaySeconds > -32);
  assert.ok(delaySeconds < -31);
  assert.deepEqual(savedDelays, [{ itemId: 'item-9', streamIndex: 0, delaySeconds }]);
});

test('preload jellyfin subtitles accepts numeric string mpv track ids', async () => {
  const commands: Array<Array<string | number>> = [];
  const preload = createPreloadJellyfinExternalSubtitlesHandler(
    makeDeps({
      listJellyfinSubtitleTracks: async () => [
        { index: 0, language: 'jpn', title: 'Japanese', deliveryUrl: 'https://sub/a.srt' },
        { index: 1, language: 'eng', title: 'English', deliveryUrl: 'https://sub/b.srt' },
      ],
      getMpvClient: () => ({
        requestProperty: async () => [
          {
            type: 'sub',
            id: ' ',
            lang: 'jpn',
            title: 'Invalid empty id',
            external: true,
            'external-filename': '/tmp/subminer-jellyfin-subtitles/invalid.srt',
          },
          {
            type: 'sub',
            id: '10',
            lang: 'jpn',
            title: 'Japanese',
            external: true,
            'external-filename': '/tmp/subminer-jellyfin-subtitles/0.srt',
          },
          {
            type: 'sub',
            id: '11',
            lang: 'eng',
            title: 'English',
            external: true,
            'external-filename': '/tmp/subminer-jellyfin-subtitles/1.srt',
          },
        ],
      }),
      sendMpvCommand: (command) => commands.push(command),
    }),
  );

  await preload({ session, clientInfo, itemId: 'item-1' });

  assert.deepEqual(setPropertyCommandsExceptTrackAutoSelection(commands), [
    ['set_property', 'sid', 10],
    ['set_property', 'secondary-sid', 11],
  ]);
});

test('preload jellyfin subtitles retries transient mpv track-list read failures', async () => {
  const commands: Array<Array<string | number>> = [];
  let requestCount = 0;
  const preload = createPreloadJellyfinExternalSubtitlesHandler(
    makeDeps({
      listJellyfinSubtitleTracks: async () => [
        { index: 0, language: 'jpn', title: 'Japanese', deliveryUrl: 'https://sub/a.srt' },
      ],
      getMpvClient: () => ({
        connected: true,
        requestProperty: async () => {
          requestCount += 1;
          if (requestCount === 1) {
            throw new Error('MPV request timed out');
          }
          return [
            {
              type: 'sub',
              id: 10,
              lang: 'jpn',
              title: 'Japanese',
              external: true,
              'external-filename': '/tmp/subminer-jellyfin-subtitles/0.srt',
            },
          ];
        },
      }),
      sendMpvCommand: (command) => commands.push(command),
    }),
  );

  await preload({ session, clientInfo, itemId: 'item-1' });

  assert.equal(requestCount, 2);
  assert.deepEqual(withoutTrackAutoSelectionCommands(commands).at(-1), ['set_property', 'sid', 10]);
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
  assert.deepEqual(setPropertyCommandsExceptTrackAutoSelection(commands), [
    ['set_property', 'sid', 11],
  ]);
});

test('preload jellyfin subtitles suppresses subtitle selection without disabling video auto selection', async () => {
  const commands: Array<Array<string | number>> = [];
  const preload = createPreloadJellyfinExternalSubtitlesHandler(
    makeDeps({
      listJellyfinSubtitleTracks: async () => [
        { index: 1, language: 'jpn', title: 'Japanese', deliveryUrl: 'https://sub/jpn.srt' },
        { index: 2, language: 'eng', title: 'English', deliveryUrl: 'https://sub/eng.srt' },
      ],
      getMpvClient: () => ({
        requestProperty: async () => [
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
            id: 12,
            lang: 'eng',
            title: 'English',
            external: true,
            'external-filename': '/tmp/subminer-jellyfin-subtitles/2.srt',
          },
        ],
      }),
      sendMpvCommand: (command) => commands.push(command),
    }),
  );

  await preload({ session, clientInfo, itemId: 'item-1' });

  const firstSubAddIndex = commands.findIndex((command) => command[0] === 'sub-add');
  const subtitleSuppressionIndex = commands.findIndex(
    (command) => command[0] === 'set_property' && command[1] === 'sid' && command[2] === 'no',
  );
  const finalPrimarySidIndex = commands.findIndex(
    (command) => command[0] === 'set_property' && command[1] === 'sid' && command[2] === 11,
  );

  assert.equal(
    commands.some(
      (command) => command[0] === 'set_property' && command[1] === 'track-auto-selection',
    ),
    false,
  );
  assert.ok(subtitleSuppressionIndex >= 0);
  assert.ok(subtitleSuppressionIndex < firstSubAddIndex);
  assert.ok(firstSubAddIndex < finalPrimarySidIndex);
  assert.equal(
    commands.filter(
      (command) => command[0] === 'set_property' && command[1] === 'sid' && command[2] === 11,
    ).length,
    1,
  );
});

test('preload jellyfin subtitles does not select a missing japanese track', async () => {
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
      (command) =>
        command[0] === 'set_property' && command[1] === 'sid' && typeof command[2] === 'number',
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
  const cleanupCalls: string[][] = [];
  const logs: string[] = [];
  let cleanupShouldFail = false;
  const preload = createPreloadJellyfinExternalSubtitlesHandler(
    makeDeps({
      listJellyfinSubtitleTracks: async (_session, _clientInfo, itemId) => [
        {
          index: itemId === 'item-1' ? 0 : 1,
          language: 'eng',
          title: 'English',
          deliveryUrl: `https://sub/${itemId}.srt`,
        },
      ],
      getMpvClient: () => ({ requestProperty: async () => [] }),
      cacheSubtitleTrack: async (track) => ({
        path: `/tmp/subminer-jellyfin-subtitles-${track.index}/track.srt`,
        cleanupDir: `/tmp/subminer-jellyfin-subtitles-${track.index}`,
      }),
      sendMpvCommand: (command) => commands.push(command),
      cleanupCachedSubtitles: (dirs) => {
        cleanupCalls.push(dirs);
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
  cleanupShouldFail = false;
  preload.cleanupCachedSubtitles();

  assert.deepEqual(logs, ['Failed to cleanup Jellyfin cached subtitles']);
  assert.deepEqual(cleanupCalls, [
    ['/tmp/subminer-jellyfin-subtitles-0'],
    ['/tmp/subminer-jellyfin-subtitles-0', '/tmp/subminer-jellyfin-subtitles-1'],
  ]);
  assert.deepEqual(
    commands.filter((command) => command[0] === 'sub-add'),
    [
      ['sub-add', '/tmp/subminer-jellyfin-subtitles-0/track.srt', 'auto', 'English', 'eng'],
      ['sub-add', '/tmp/subminer-jellyfin-subtitles-1/track.srt', 'auto', 'English', 'eng'],
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

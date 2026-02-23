import test from 'node:test';
import assert from 'node:assert/strict';
import { createHandleJellyfinListCommands } from './jellyfin-cli-list';

const baseSession = {
  serverUrl: 'http://localhost',
  accessToken: 'token',
  userId: 'user-id',
  username: 'user',
};

const baseClientInfo = {
  clientName: 'SubMiner',
  clientVersion: '1.0.0',
  deviceId: 'device-id',
};

const baseConfig = {
  defaultLibraryId: '',
};

test('list handler no-ops when no list command is set', async () => {
  const handler = createHandleJellyfinListCommands({
    listJellyfinLibraries: async () => [],
    listJellyfinItems: async () => [],
    listJellyfinSubtitleTracks: async () => [],
    logInfo: () => {},
  });

  const handled = await handler({
    args: {
      jellyfinLibraries: false,
      jellyfinItems: false,
      jellyfinSubtitles: false,
    } as never,
    session: baseSession,
    clientInfo: baseClientInfo,
    jellyfinConfig: baseConfig,
  });

  assert.equal(handled, false);
});

test('list handler logs libraries', async () => {
  const logs: string[] = [];
  const handler = createHandleJellyfinListCommands({
    listJellyfinLibraries: async () => [{ id: 'lib1', name: 'Anime', collectionType: 'tvshows' }],
    listJellyfinItems: async () => [],
    listJellyfinSubtitleTracks: async () => [],
    logInfo: (message) => logs.push(message),
  });

  const handled = await handler({
    args: {
      jellyfinLibraries: true,
      jellyfinItems: false,
      jellyfinSubtitles: false,
    } as never,
    session: baseSession,
    clientInfo: baseClientInfo,
    jellyfinConfig: baseConfig,
  });

  assert.equal(handled, true);
  assert.ok(logs.some((line) => line.includes('Jellyfin library: Anime [lib1] (tvshows)')));
});

test('list handler resolves items using default library id', async () => {
  let usedLibraryId = '';
  const logs: string[] = [];
  const handler = createHandleJellyfinListCommands({
    listJellyfinLibraries: async () => [],
    listJellyfinItems: async (_session, _clientInfo, params) => {
      usedLibraryId = params.libraryId;
      return [{ id: 'item1', title: 'Episode 1', type: 'Episode' }];
    },
    listJellyfinSubtitleTracks: async () => [],
    logInfo: (message) => logs.push(message),
  });

  const handled = await handler({
    args: {
      jellyfinLibraries: false,
      jellyfinItems: true,
      jellyfinSubtitles: false,
      jellyfinLibraryId: '',
      jellyfinSearch: 'episode',
      jellyfinLimit: 10,
    } as never,
    session: baseSession,
    clientInfo: baseClientInfo,
    jellyfinConfig: {
      defaultLibraryId: 'default-lib',
    },
  });

  assert.equal(handled, true);
  assert.equal(usedLibraryId, 'default-lib');
  assert.ok(logs.some((line) => line.includes('Jellyfin item: Episode 1 [item1] (Episode)')));
});

test('list handler throws when items command has no library id', async () => {
  const handler = createHandleJellyfinListCommands({
    listJellyfinLibraries: async () => [],
    listJellyfinItems: async () => [],
    listJellyfinSubtitleTracks: async () => [],
    logInfo: () => {},
  });

  await assert.rejects(
    handler({
      args: {
        jellyfinLibraries: false,
        jellyfinItems: true,
        jellyfinSubtitles: false,
        jellyfinLibraryId: '',
      } as never,
      session: baseSession,
      clientInfo: baseClientInfo,
      jellyfinConfig: baseConfig,
    }),
    /Missing Jellyfin library id/,
  );
});

test('list handler logs subtitle urls only when requested', async () => {
  const logs: string[] = [];
  const handler = createHandleJellyfinListCommands({
    listJellyfinLibraries: async () => [],
    listJellyfinItems: async () => [],
    listJellyfinSubtitleTracks: async () => [
      { index: 1, deliveryUrl: 'http://localhost/sub1.srt', language: 'eng' },
      { index: 2, language: 'jpn' },
    ],
    logInfo: (message) => logs.push(message),
  });

  const handled = await handler({
    args: {
      jellyfinLibraries: false,
      jellyfinItems: false,
      jellyfinSubtitles: true,
      jellyfinItemId: 'item1',
      jellyfinSubtitleUrlsOnly: true,
    } as never,
    session: baseSession,
    clientInfo: baseClientInfo,
    jellyfinConfig: baseConfig,
  });

  assert.equal(handled, true);
  assert.deepEqual(logs, ['http://localhost/sub1.srt']);
});

test('list handler throws when subtitle command has no item id', async () => {
  const handler = createHandleJellyfinListCommands({
    listJellyfinLibraries: async () => [],
    listJellyfinItems: async () => [],
    listJellyfinSubtitleTracks: async () => [],
    logInfo: () => {},
  });

  await assert.rejects(
    handler({
      args: {
        jellyfinLibraries: false,
        jellyfinItems: false,
        jellyfinSubtitles: true,
      } as never,
      session: baseSession,
      clientInfo: baseClientInfo,
      jellyfinConfig: baseConfig,
    }),
    /Missing --jellyfin-item-id/,
  );
});

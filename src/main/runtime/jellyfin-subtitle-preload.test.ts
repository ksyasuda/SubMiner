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

test('preload jellyfin subtitles adds external tracks and chooses japanese+english tracks', async () => {
  const commands: Array<Array<string | number>> = [];
  const preload = createPreloadJellyfinExternalSubtitlesHandler({
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
    wait: async () => {},
    logDebug: () => {},
  });

  await preload({ session, clientInfo, itemId: 'item-1' });

  assert.deepEqual(commands, [
    ['sub-add', 'https://sub/a.srt', 'cached', 'Japanese', 'jpn'],
    ['sub-add', 'https://sub/b.srt', 'cached', 'English SDH', 'eng'],
    ['set_property', 'sid', 5],
    ['set_property', 'secondary-sid', 6],
  ]);
});

test('preload jellyfin subtitles exits quietly when no external tracks', async () => {
  const commands: Array<Array<string | number>> = [];
  let waited = false;
  const preload = createPreloadJellyfinExternalSubtitlesHandler({
    listJellyfinSubtitleTracks: async () => [{ index: 0, language: 'jpn', title: 'Embedded' }],
    getMpvClient: () => ({ requestProperty: async () => [] }),
    sendMpvCommand: (command) => commands.push(command),
    wait: async () => {
      waited = true;
    },
    logDebug: () => {},
  });

  await preload({ session, clientInfo, itemId: 'item-1' });

  assert.equal(waited, false);
  assert.deepEqual(commands, []);
});

test('preload jellyfin subtitles logs debug on failure', async () => {
  const logs: string[] = [];
  const preload = createPreloadJellyfinExternalSubtitlesHandler({
    listJellyfinSubtitleTracks: async () => {
      throw new Error('network down');
    },
    getMpvClient: () => null,
    sendMpvCommand: () => {},
    wait: async () => {},
    logDebug: (message) => logs.push(message),
  });

  await preload({ session, clientInfo, itemId: 'item-1' });

  assert.deepEqual(logs, ['Failed to preload Jellyfin external subtitles']);
});

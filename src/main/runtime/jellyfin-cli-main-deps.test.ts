import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildHandleJellyfinAuthCommandsMainDepsHandler,
  createBuildHandleJellyfinListCommandsMainDepsHandler,
  createBuildHandleJellyfinPlayCommandMainDepsHandler,
  createBuildHandleJellyfinRemoteAnnounceCommandMainDepsHandler,
} from './jellyfin-cli-main-deps';

test('jellyfin auth commands main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const deps = createBuildHandleJellyfinAuthCommandsMainDepsHandler({
    patchRawConfig: () => calls.push('patch'),
    authenticateWithPassword: async () => ({}) as never,
    logInfo: (message) => calls.push(`info:${message}`),
  })();

  deps.patchRawConfig({ jellyfin: {} });
  await deps.authenticateWithPassword('', '', '', {
    deviceId: '',
    clientName: '',
    clientVersion: '',
  });
  deps.logInfo('ok');
  assert.deepEqual(calls, ['patch', 'info:ok']);
});

test('jellyfin list commands main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const deps = createBuildHandleJellyfinListCommandsMainDepsHandler({
    listJellyfinLibraries: async () => {
      calls.push('libraries');
      return [];
    },
    listJellyfinItems: async () => {
      calls.push('items');
      return [];
    },
    listJellyfinSubtitleTracks: async () => {
      calls.push('subtitles');
      return [];
    },
    logInfo: (message) => calls.push(`info:${message}`),
  })();

  await deps.listJellyfinLibraries({} as never, {} as never);
  await deps.listJellyfinItems({} as never, {} as never, { libraryId: '', limit: 1 });
  await deps.listJellyfinSubtitleTracks({} as never, {} as never, 'id');
  deps.logInfo('done');
  assert.deepEqual(calls, ['libraries', 'items', 'subtitles', 'info:done']);
});

test('jellyfin play command main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const deps = createBuildHandleJellyfinPlayCommandMainDepsHandler({
    playJellyfinItemInMpv: async () => {
      calls.push('play');
    },
    logWarn: (message) => calls.push(`warn:${message}`),
  })();

  await deps.playJellyfinItemInMpv({} as never);
  deps.logWarn('missing');
  assert.deepEqual(calls, ['play', 'warn:missing']);
});

test('jellyfin remote announce main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const session = { advertiseNow: async () => true };
  const deps = createBuildHandleJellyfinRemoteAnnounceCommandMainDepsHandler({
    startJellyfinRemoteSession: async () => {
      calls.push('start');
    },
    getRemoteSession: () => session,
    logInfo: (message) => calls.push(`info:${message}`),
    logWarn: (message) => calls.push(`warn:${message}`),
  })();

  await deps.startJellyfinRemoteSession();
  assert.equal(deps.getRemoteSession(), session);
  deps.logInfo('visible');
  deps.logWarn('not-visible');
  assert.deepEqual(calls, ['start', 'info:visible', 'warn:not-visible']);
});

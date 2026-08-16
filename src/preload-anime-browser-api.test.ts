import assert from 'node:assert/strict';
import test from 'node:test';
import { IPC_CHANNELS } from './shared/ipc/contracts';
import { createAnimeBrowserAPI } from './preload-anime-browser-api';

function createIpcRecorder() {
  const calls: Array<{ channel: string; args: unknown[] }> = [];
  return {
    calls,
    ipc: {
      invoke: async (channel: string, ...args: unknown[]) => {
        calls.push({ channel, args });
        return undefined;
      },
      on: () => {},
      removeListener: () => {},
    },
  };
}

test('anime browser API keeps one opaque session per renderer bridge', async () => {
  const first = createIpcRecorder();
  const second = createIpcRecorder();
  const firstApi = createAnimeBrowserAPI(first.ipc as never);
  const secondApi = createAnimeBrowserAPI(second.ipc as never);

  await firstApi.getSnapshot();
  await firstApi.selectSource('source.one');
  await firstApi.search('frieren', 2);
  await firstApi.getDetails('/frieren');
  await firstApi.getEpisodes('/frieren', 'source.one');
  await firstApi.getPlaybackState();
  await secondApi.getPopular(3);

  const firstSessionId = first.calls[0]?.args[0];
  const secondSessionId = second.calls[0]?.args[0];
  assert.equal(typeof firstSessionId, 'string');
  assert.notEqual(firstSessionId, secondSessionId);
  assert.deepEqual(first.calls, [
    {
      channel: IPC_CHANNELS.request.animeBrowserGetSnapshot,
      args: [firstSessionId],
    },
    {
      channel: IPC_CHANNELS.request.animeBrowserSelectSource,
      args: [firstSessionId, 'source.one'],
    },
    {
      channel: IPC_CHANNELS.request.animeBrowserSearch,
      args: [firstSessionId, 'frieren', 2],
    },
    {
      channel: IPC_CHANNELS.request.animeBrowserGetDetails,
      args: [firstSessionId, '/frieren', undefined],
    },
    {
      channel: IPC_CHANNELS.request.animeBrowserGetEpisodes,
      args: [firstSessionId, '/frieren', 'source.one'],
    },
    {
      channel: IPC_CHANNELS.request.animeBrowserGetPlaybackState,
      args: [],
    },
  ]);
  assert.deepEqual(second.calls, [
    {
      channel: IPC_CHANNELS.request.animeBrowserGetPopular,
      args: [secondSessionId, 3],
    },
  ]);
});

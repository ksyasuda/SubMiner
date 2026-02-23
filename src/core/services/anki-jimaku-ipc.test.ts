import assert from 'node:assert/strict';
import test from 'node:test';
import { registerAnkiJimakuIpcHandlers } from './anki-jimaku-ipc';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';

function createFakeRegistrar(): {
  registrar: {
    on: (channel: string, listener: (event: unknown, ...args: unknown[]) => void) => void;
    handle: (channel: string, listener: (event: unknown, ...args: unknown[]) => unknown) => void;
  };
  onHandlers: Map<string, (event: unknown, ...args: unknown[]) => void>;
  handleHandlers: Map<string, (event: unknown, ...args: unknown[]) => unknown>;
} {
  const onHandlers = new Map<string, (event: unknown, ...args: unknown[]) => void>();
  const handleHandlers = new Map<string, (event: unknown, ...args: unknown[]) => unknown>();
  return {
    registrar: {
      on: (channel, listener) => {
        onHandlers.set(channel, listener);
      },
      handle: (channel, listener) => {
        handleHandlers.set(channel, listener);
      },
    },
    onHandlers,
    handleHandlers,
  };
}

test('anki/jimaku IPC handlers reject malformed invoke payloads', async () => {
  const { registrar, handleHandlers } = createFakeRegistrar();
  let previewCalls = 0;
  registerAnkiJimakuIpcHandlers(
    {
      setAnkiConnectEnabled: () => {},
      clearAnkiHistory: () => {},
      refreshKnownWords: async () => {},
      respondFieldGrouping: () => {},
      buildKikuMergePreview: async () => {
        previewCalls += 1;
        return { ok: true };
      },
      getJimakuMediaInfo: () => ({
        title: 'x',
        season: null,
        episode: null,
        confidence: 'high',
        filename: 'x.mkv',
        rawTitle: 'x',
      }),
      searchJimakuEntries: async () => ({ ok: true, data: [] }),
      listJimakuFiles: async () => ({ ok: true, data: [] }),
      resolveJimakuApiKey: async () => 'token',
      getCurrentMediaPath: () => '/tmp/a.mkv',
      isRemoteMediaPath: () => false,
      downloadToFile: async () => ({ ok: true, path: '/tmp/sub.ass' }),
      onDownloadedSubtitle: () => {},
    },
    registrar,
  );

  const previewHandler = handleHandlers.get(IPC_CHANNELS.request.kikuBuildMergePreview);
  assert.ok(previewHandler);
  const invalidPreviewResult = await previewHandler!({}, null);
  assert.deepEqual(invalidPreviewResult, {
    ok: false,
    error: 'Invalid merge preview request payload',
  });
  await previewHandler!({}, { keepNoteId: 1, deleteNoteId: 2, deleteDuplicate: false });
  assert.equal(previewCalls, 1);

  const searchHandler = handleHandlers.get(IPC_CHANNELS.request.jimakuSearchEntries);
  assert.ok(searchHandler);
  const invalidSearchResult = await searchHandler!({}, { query: 12 });
  assert.deepEqual(invalidSearchResult, {
    ok: false,
    error: { error: 'Invalid Jimaku search query payload', code: 400 },
  });

  const filesHandler = handleHandlers.get(IPC_CHANNELS.request.jimakuListFiles);
  assert.ok(filesHandler);
  const invalidFilesResult = await filesHandler!({}, { entryId: 'x' });
  assert.deepEqual(invalidFilesResult, {
    ok: false,
    error: { error: 'Invalid Jimaku files query payload', code: 400 },
  });

  const downloadHandler = handleHandlers.get(IPC_CHANNELS.request.jimakuDownloadFile);
  assert.ok(downloadHandler);
  const invalidDownloadResult = await downloadHandler!({}, { entryId: 1, url: '/x' });
  assert.deepEqual(invalidDownloadResult, {
    ok: false,
    error: { error: 'Invalid Jimaku download query payload', code: 400 },
  });
});

test('anki/jimaku IPC command handlers ignore malformed payloads', () => {
  const { registrar, onHandlers } = createFakeRegistrar();
  const fieldGroupingChoices: unknown[] = [];
  const enabledStates: boolean[] = [];
  registerAnkiJimakuIpcHandlers(
    {
      setAnkiConnectEnabled: (enabled) => {
        enabledStates.push(enabled);
      },
      clearAnkiHistory: () => {},
      refreshKnownWords: async () => {},
      respondFieldGrouping: (choice) => {
        fieldGroupingChoices.push(choice);
      },
      buildKikuMergePreview: async () => ({ ok: true }),
      getJimakuMediaInfo: () => ({
        title: 'x',
        season: null,
        episode: null,
        confidence: 'high',
        filename: 'x.mkv',
        rawTitle: 'x',
      }),
      searchJimakuEntries: async () => ({ ok: true, data: [] }),
      listJimakuFiles: async () => ({ ok: true, data: [] }),
      resolveJimakuApiKey: async () => 'token',
      getCurrentMediaPath: () => '/tmp/a.mkv',
      isRemoteMediaPath: () => false,
      downloadToFile: async () => ({ ok: true, path: '/tmp/sub.ass' }),
      onDownloadedSubtitle: () => {},
    },
    registrar,
  );

  onHandlers.get(IPC_CHANNELS.command.setAnkiConnectEnabled)!({}, 'true');
  onHandlers.get(IPC_CHANNELS.command.setAnkiConnectEnabled)!({}, true);
  assert.deepEqual(enabledStates, [true]);

  onHandlers.get(IPC_CHANNELS.command.kikuFieldGroupingRespond)!({}, null);
  onHandlers.get(IPC_CHANNELS.command.kikuFieldGroupingRespond)!(
    {},
    {
      keepNoteId: 1,
      deleteNoteId: 2,
      deleteDuplicate: false,
      cancelled: false,
    },
  );
  assert.deepEqual(fieldGroupingChoices, [
    {
      keepNoteId: 1,
      deleteNoteId: 2,
      deleteDuplicate: false,
      cancelled: false,
    },
  ]);
});

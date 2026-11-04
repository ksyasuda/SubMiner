import assert from 'node:assert/strict';
import test from 'node:test';

import type { ElectronAPI } from '../../types';
import { createRendererState } from '../state.js';
import { createJimakuModal } from './jimaku.js';

function createClassList(initialTokens: string[] = []) {
  const tokens = new Set(initialTokens);
  return {
    add: (...entries: string[]) => {
      for (const entry of entries) {
        tokens.add(entry);
      }
    },
    remove: (...entries: string[]) => {
      for (const entry of entries) {
        tokens.delete(entry);
      }
    },
    contains: (entry: string) => tokens.has(entry),
  };
}

function createElementStub() {
  const classList = createClassList();
  return {
    textContent: '',
    className: '',
    style: {},
    classList,
    children: [] as unknown[],
    appendChild(child: unknown) {
      this.children.push(child);
    },
    addEventListener: () => {},
  };
}

function createListStub() {
  return {
    innerHTML: '',
    children: [] as unknown[],
    appendChild(child: unknown) {
      this.children.push(child);
    },
  };
}

function flushAsyncWork(): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, 0);
  });
}

test('successful Jimaku subtitle selection closes modal', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;

  const modalCloseNotifications: Array<'runtime-options' | 'subsync' | 'jimaku' | 'kiku'> = [];

  const electronAPI = {
    jimakuDownloadFile: async () => ({ ok: true, path: '/tmp/subtitles/episode01.ass' }),
    notifyOverlayModalClosed: (modal: 'runtime-options' | 'subsync' | 'jimaku' | 'kiku') => {
      modalCloseNotifications.push(modal);
    },
  } as unknown as ElectronAPI;

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: { electronAPI },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      activeElement: null,
      createElement: () => createElementStub(),
    },
  });

  try {
    const overlayClassList = createClassList(['interactive']);
    const jimakuModalClassList = createClassList();
    const jimakuEntriesSectionClassList = createClassList(['hidden']);
    const jimakuFilesSectionClassList = createClassList();
    const jimakuBroadenButtonClassList = createClassList(['hidden']);
    const state = createRendererState();
    state.jimakuModalOpen = true;
    state.currentEntryId = 42;
    state.selectedFileIndex = 0;
    state.jimakuFiles = [
      {
        name: 'episode01.ass',
        url: 'https://jimaku.cc/files/episode01.ass',
        size: 1000,
        last_modified: '2026-03-01',
      },
    ];

    const ctx = {
      dom: {
        overlay: { classList: overlayClassList },
        jimakuModal: {
          classList: jimakuModalClassList,
          setAttribute: () => {},
        },
        jimakuTitleInput: { value: '' },
        jimakuSeasonInput: { value: '' },
        jimakuEpisodeInput: { value: '' },
        jimakuSearchButton: { addEventListener: () => {} },
        jimakuCloseButton: { addEventListener: () => {} },
        jimakuStatus: { textContent: '', style: { color: '' } },
        jimakuEntriesSection: { classList: jimakuEntriesSectionClassList },
        jimakuEntriesList: createListStub(),
        jimakuFilesSection: { classList: jimakuFilesSectionClassList },
        jimakuFilesList: createListStub(),
        jimakuBroadenButton: {
          classList: jimakuBroadenButtonClassList,
          addEventListener: () => {},
        },
        jimakuTabAnimeButton: { classList: createClassList(['active']), setAttribute: () => {} },
        jimakuTabLiveActionButton: { classList: createClassList(), setAttribute: () => {} },
      },
      state,
    };

    const jimakuModal = createJimakuModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
    });

    let prevented = false;
    jimakuModal.handleJimakuKeydown({
      key: 'Enter',
      preventDefault: () => {
        prevented = true;
      },
    } as KeyboardEvent);
    await flushAsyncWork();

    assert.equal(prevented, true);
    assert.equal(state.jimakuModalOpen, false);
    assert.equal(jimakuModalClassList.contains('hidden'), true);
    assert.equal(overlayClassList.contains('interactive'), false);
    assert.deepEqual(modalCloseNotifications, ['jimaku']);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('switching to the Live action tab re-runs the search with the live action category', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;

  const searchQueries: Array<{ query: string; category?: string }> = [];
  const electronAPI = {
    jimakuSearchEntries: async (query: { query: string; category?: string }) => {
      searchQueries.push(query);
      return { ok: true, data: [] };
    },
  } as unknown as ElectronAPI;

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: { electronAPI },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      activeElement: null,
      createElement: () => createElementStub(),
    },
  });

  try {
    const state = createRendererState();
    state.jimakuModalOpen = true;
    const animeTabClassList = createClassList(['active']);
    const liveActionTabClassList = createClassList();
    const status = { textContent: '', style: { color: '' } };

    const ctx = {
      dom: {
        overlay: { classList: createClassList(['interactive']) },
        jimakuModal: { classList: createClassList(), setAttribute: () => {} },
        jimakuTitleInput: { value: 'Shinzanmono' },
        jimakuSeasonInput: { value: '' },
        jimakuEpisodeInput: { value: '3' },
        jimakuSearchButton: { addEventListener: () => {} },
        jimakuCloseButton: { addEventListener: () => {} },
        jimakuStatus: status,
        jimakuEntriesSection: { classList: createClassList(['hidden']) },
        jimakuEntriesList: createListStub(),
        jimakuFilesSection: { classList: createClassList(['hidden']) },
        jimakuFilesList: createListStub(),
        jimakuBroadenButton: { classList: createClassList(['hidden']), addEventListener: () => {} },
        jimakuTabAnimeButton: { classList: animeTabClassList, setAttribute: () => {} },
        jimakuTabLiveActionButton: { classList: liveActionTabClassList, setAttribute: () => {} },
      },
      state,
    };

    const jimakuModal = createJimakuModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
    });

    jimakuModal.handleJimakuKeydown({
      key: 'ArrowRight',
      preventDefault: () => {},
    } as KeyboardEvent);
    await flushAsyncWork();

    assert.equal(state.jimakuActiveTab, 'liveAction');
    assert.equal(liveActionTabClassList.contains('active'), true);
    assert.equal(animeTabClassList.contains('active'), false);
    assert.deepEqual(searchQueries, [{ query: 'Shinzanmono', category: 'liveAction' }]);
    assert.equal(status.textContent, 'No live action entries found. Try the Anime tab.');

    // Same tab again is a no-op: no duplicate request.
    jimakuModal.handleJimakuKeydown({
      key: 'ArrowRight',
      preventDefault: () => {},
    } as KeyboardEvent);
    await flushAsyncWork();
    assert.equal(searchQueries.length, 1);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('a slow reply from a superseded search does not overwrite the newer results', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;

  const pending: Array<(entries: unknown[]) => void> = [];
  const electronAPI = {
    jimakuSearchEntries: () =>
      new Promise((resolve) => {
        pending.push((entries) => resolve({ ok: true, data: entries }));
      }),
    jimakuListFiles: async () => ({ ok: true, data: [] }),
  } as unknown as ElectronAPI;

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: { electronAPI },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      activeElement: null,
      createElement: () => createElementStub(),
    },
  });

  try {
    const state = createRendererState();
    state.jimakuModalOpen = true;

    const ctx = {
      dom: {
        overlay: { classList: createClassList(['interactive']) },
        jimakuModal: { classList: createClassList(), setAttribute: () => {} },
        jimakuTitleInput: { value: 'Shinzanmono' },
        jimakuSeasonInput: { value: '' },
        jimakuEpisodeInput: { value: '' },
        jimakuSearchButton: { addEventListener: () => {} },
        jimakuCloseButton: { addEventListener: () => {} },
        jimakuStatus: { textContent: '', style: { color: '' } },
        jimakuEntriesSection: { classList: createClassList(['hidden']) },
        jimakuEntriesList: createListStub(),
        jimakuFilesSection: { classList: createClassList(['hidden']) },
        jimakuFilesList: createListStub(),
        jimakuBroadenButton: { classList: createClassList(['hidden']), addEventListener: () => {} },
        jimakuTabAnimeButton: { classList: createClassList(['active']), setAttribute: () => {} },
        jimakuTabLiveActionButton: { classList: createClassList(), setAttribute: () => {} },
      },
      state,
    };

    const jimakuModal = createJimakuModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
    });

    // Anime -> Live action -> Anime, all before any reply arrives.
    jimakuModal.handleJimakuKeydown({
      key: 'ArrowRight',
      preventDefault: () => {},
    } as KeyboardEvent);
    jimakuModal.handleJimakuKeydown({
      key: 'ArrowLeft',
      preventDefault: () => {},
    } as KeyboardEvent);
    await flushAsyncWork();
    assert.equal(pending.length, 2);

    // The stale live action reply lands after the newer anime search was issued.
    pending[0]!([{ id: 1, name: 'Stale live action entry' }]);
    await flushAsyncWork();
    assert.equal(state.jimakuEntries.length, 0);

    pending[1]!([
      { id: 2, name: 'Anime A' },
      { id: 3, name: 'Anime B' },
    ]);
    await flushAsyncWork();
    assert.deepEqual(
      state.jimakuEntries.map((entry) => entry.id),
      [2, 3],
    );
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('closing the modal discards an in-flight search reply', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;

  let resolveSearch!: (entries: unknown[]) => void;
  let listFilesCalls = 0;
  const electronAPI = {
    jimakuSearchEntries: () =>
      new Promise((resolve) => {
        resolveSearch = (entries) => resolve({ ok: true, data: entries });
      }),
    jimakuListFiles: async () => {
      listFilesCalls += 1;
      return { ok: true, data: [] };
    },
    notifyOverlayModalClosed: () => {},
  } as unknown as ElectronAPI;

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: { electronAPI },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      activeElement: null,
      createElement: () => createElementStub(),
    },
  });

  try {
    const state = createRendererState();
    state.jimakuModalOpen = true;

    const ctx = {
      dom: {
        overlay: { classList: createClassList(['interactive']) },
        jimakuModal: { classList: createClassList(), setAttribute: () => {} },
        jimakuTitleInput: { value: 'Shinzanmono' },
        jimakuSeasonInput: { value: '' },
        jimakuEpisodeInput: { value: '' },
        jimakuSearchButton: { addEventListener: () => {} },
        jimakuCloseButton: { addEventListener: () => {} },
        jimakuStatus: { textContent: '', style: { color: '' } },
        jimakuEntriesSection: { classList: createClassList(['hidden']) },
        jimakuEntriesList: createListStub(),
        jimakuFilesSection: { classList: createClassList(['hidden']) },
        jimakuFilesList: createListStub(),
        jimakuBroadenButton: { classList: createClassList(['hidden']), addEventListener: () => {} },
        jimakuTabAnimeButton: { classList: createClassList(['active']), setAttribute: () => {} },
        jimakuTabLiveActionButton: { classList: createClassList(), setAttribute: () => {} },
      },
      state,
    };

    const jimakuModal = createJimakuModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
    });

    jimakuModal.handleJimakuKeydown({ key: 'Enter', preventDefault: () => {} } as KeyboardEvent);
    await flushAsyncWork();
    jimakuModal.closeJimakuModal();

    // A single entry would normally auto-select and fetch its files.
    resolveSearch([{ id: 7, name: 'Only entry' }]);
    await flushAsyncWork();

    assert.equal(state.jimakuEntries.length, 0);
    assert.equal(state.currentEntryId, null);
    assert.equal(listFilesCalls, 0);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('a slow files reply for a previously selected entry is ignored', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;

  const pending = new Map<number, (files: unknown[]) => void>();
  const electronAPI = {
    jimakuListFiles: (query: { entryId: number }) =>
      new Promise((resolve) => {
        pending.set(query.entryId, (files) => resolve({ ok: true, data: files }));
      }),
  } as unknown as ElectronAPI;

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: { electronAPI },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      activeElement: null,
      createElement: () => createElementStub(),
    },
  });

  try {
    const state = createRendererState();
    state.jimakuModalOpen = true;
    state.jimakuEntries = [
      { id: 1, name: 'Entry A' },
      { id: 2, name: 'Entry B' },
    ];

    const ctx = {
      dom: {
        overlay: { classList: createClassList(['interactive']) },
        jimakuModal: { classList: createClassList(), setAttribute: () => {} },
        jimakuTitleInput: { value: '' },
        jimakuSeasonInput: { value: '' },
        jimakuEpisodeInput: { value: '' },
        jimakuSearchButton: { addEventListener: () => {} },
        jimakuCloseButton: { addEventListener: () => {} },
        jimakuStatus: { textContent: '', style: { color: '' } },
        jimakuEntriesSection: { classList: createClassList() },
        jimakuEntriesList: createListStub(),
        jimakuFilesSection: { classList: createClassList(['hidden']) },
        jimakuFilesList: createListStub(),
        jimakuBroadenButton: { classList: createClassList(['hidden']), addEventListener: () => {} },
        jimakuTabAnimeButton: { classList: createClassList(['active']), setAttribute: () => {} },
        jimakuTabLiveActionButton: { classList: createClassList(), setAttribute: () => {} },
      },
      state,
    };

    const jimakuModal = createJimakuModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
    });

    // Select entry A, then move to entry B before A's files arrive.
    jimakuModal.handleJimakuKeydown({ key: 'Enter', preventDefault: () => {} } as KeyboardEvent);
    jimakuModal.handleJimakuKeydown({
      key: 'ArrowDown',
      preventDefault: () => {},
    } as KeyboardEvent);
    jimakuModal.handleJimakuKeydown({ key: 'Enter', preventDefault: () => {} } as KeyboardEvent);
    await flushAsyncWork();
    assert.equal(state.currentEntryId, 2);

    pending.get(1)!([
      { name: 'a.srt', url: 'https://jimaku.cc/a.srt', size: 1, last_modified: '' },
    ]);
    await flushAsyncWork();
    assert.equal(state.jimakuFiles.length, 0);

    pending.get(2)!([
      { name: 'b1.srt', url: 'https://jimaku.cc/b1.srt', size: 1, last_modified: '' },
      { name: 'b2.srt', url: 'https://jimaku.cc/b2.srt', size: 1, last_modified: '' },
    ]);
    await flushAsyncWork();
    assert.deepEqual(
      state.jimakuFiles.map((file) => file.name),
      ['b1.srt', 'b2.srt'],
    );
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('media info arriving after the modal closed does not fill inputs or search', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;

  let resolveMediaInfo!: (info: unknown) => void;
  let searchCalls = 0;
  const electronAPI = {
    getJimakuMediaInfo: () =>
      new Promise((resolve) => {
        resolveMediaInfo = resolve;
      }),
    jimakuSearchEntries: async () => {
      searchCalls += 1;
      return { ok: true, data: [] };
    },
    notifyOverlayModalClosed: () => {},
  } as unknown as ElectronAPI;

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: { electronAPI },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      activeElement: null,
      createElement: () => createElementStub(),
    },
  });

  try {
    const state = createRendererState();
    const titleInput = { value: '' };
    const status = { textContent: '', style: { color: '' } };

    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        jimakuModal: { classList: createClassList(['hidden']), setAttribute: () => {} },
        jimakuTitleInput: titleInput,
        jimakuSeasonInput: { value: '' },
        jimakuEpisodeInput: { value: '' },
        jimakuSearchButton: { addEventListener: () => {} },
        jimakuCloseButton: { addEventListener: () => {} },
        jimakuStatus: status,
        jimakuEntriesSection: { classList: createClassList(['hidden']) },
        jimakuEntriesList: createListStub(),
        jimakuFilesSection: { classList: createClassList(['hidden']) },
        jimakuFilesList: createListStub(),
        jimakuBroadenButton: { classList: createClassList(['hidden']), addEventListener: () => {} },
        jimakuTabAnimeButton: { classList: createClassList(['active']), setAttribute: () => {} },
        jimakuTabLiveActionButton: { classList: createClassList(), setAttribute: () => {} },
      },
      state,
    };

    const jimakuModal = createJimakuModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
    });

    jimakuModal.openJimakuModal();
    await flushAsyncWork();
    jimakuModal.closeJimakuModal();

    resolveMediaInfo({
      title: 'Shinzanmono',
      season: 1,
      episode: 3,
      confidence: 'high',
      filename: 'Shinzanmono S01E03.mkv',
      rawTitle: 'Shinzanmono S01E03',
    });
    await flushAsyncWork();

    assert.equal(titleInput.value, '');
    assert.equal(searchCalls, 0);
    assert.equal(status.textContent, 'Loading media info...');
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

type MediaInfoFixture = { filename: string; season: number; episode: number };

const mediaInfo = (info: MediaInfoFixture) => ({
  ...info,
  title: 'Show',
  confidence: 'medium',
  rawTitle: info.filename,
});

// A Jimaku modal whose media-info replies are either immediate (`current`) or held until a test
// resolves them (`deferReplies`), to drive re-open and reply-ordering races.
async function withReopenHarness(
  run: (h: {
    modal: ReturnType<typeof createJimakuModal>;
    episodeInput: { value: string };
    setCurrent: (info: MediaInfoFixture) => void;
    deferReplies: () => void;
    pendingReplies: Array<(info: MediaInfoFixture) => void>;
    pressEnter: () => void;
  }) => Promise<void>,
): Promise<void> {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;

  let current = mediaInfo({ filename: '01.mkv', season: 1, episode: 1 });
  let deferred = false;
  const pendingReplies: Array<(info: MediaInfoFixture) => void> = [];
  const electronAPI = {
    getJimakuMediaInfo: () =>
      deferred
        ? new Promise((resolve) => pendingReplies.push((info) => resolve(mediaInfo(info))))
        : Promise.resolve(current),
    jimakuSearchEntries: async () => ({ ok: true, data: [] }),
    notifyOverlayModalClosed: () => {},
  } as unknown as ElectronAPI;
  Object.defineProperty(globalThis, 'window', { configurable: true, value: { electronAPI } });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: { activeElement: null, createElement: () => createElementStub() },
  });

  try {
    const episodeInput = { value: '' };
    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        jimakuModal: { classList: createClassList(['hidden']), setAttribute: () => {} },
        jimakuTitleInput: { value: '' },
        jimakuSeasonInput: { value: '' },
        jimakuEpisodeInput: episodeInput,
        jimakuSearchButton: { addEventListener: () => {} },
        jimakuCloseButton: { addEventListener: () => {} },
        jimakuStatus: { textContent: '', style: { color: '' } },
        jimakuEntriesSection: { classList: createClassList(['hidden']) },
        jimakuEntriesList: createListStub(),
        jimakuFilesSection: { classList: createClassList(['hidden']) },
        jimakuFilesList: createListStub(),
        jimakuBroadenButton: { classList: createClassList(['hidden']), addEventListener: () => {} },
        jimakuTabAnimeButton: { classList: createClassList(['active']), setAttribute: () => {} },
        jimakuTabLiveActionButton: { classList: createClassList(), setAttribute: () => {} },
      },
      state: createRendererState(),
    };
    const modal = createJimakuModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
    });

    await run({
      modal,
      episodeInput,
      setCurrent: (info) => {
        current = mediaInfo(info);
      },
      deferReplies: () => {
        deferred = true;
      },
      pendingReplies,
      // Enter with no results and no focused input starts a search with the current inputs.
      pressEnter: () => {
        modal.handleJimakuKeydown({ key: 'Enter', preventDefault: () => {} } as KeyboardEvent);
      },
    });
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
}

for (const c of [
  {
    name: 'reloads media info when the video changed',
    next: { filename: '02.mkv', season: 1, episode: 2 },
    reloads: true,
  },
  {
    name: 'reloads media info for the same file name in another season',
    next: { filename: '01.mkv', season: 2, episode: 1 },
    reloads: true,
  },
  {
    name: 'keeps the current search for the same video',
    next: { filename: '01.mkv', season: 1, episode: 1 },
    reloads: false,
  },
]) {
  test(`re-opening the open Jimaku modal ${c.name}`, async () => {
    await withReopenHarness(async (h) => {
      h.modal.openJimakuModal();
      await flushAsyncWork();
      h.episodeInput.value = '7'; // the user's own edit
      h.setCurrent(c.next);
      h.modal.openJimakuModal();
      await flushAsyncWork();
      await flushAsyncWork();

      assert.equal(h.episodeInput.value, c.reloads ? String(c.next.episode) : '7');
    });
  });
}

test('a search started during a changed-video refresh does not cancel the refresh', async () => {
  await withReopenHarness(async (h) => {
    h.modal.openJimakuModal();
    await flushAsyncWork();
    h.deferReplies();

    h.modal.openJimakuModal(); // refresh request now pending
    h.pressEnter();
    await flushAsyncWork();
    h.pendingReplies.shift()?.({ filename: '02.mkv', season: 1, episode: 2 });
    await flushAsyncWork();
    h.pendingReplies.shift()?.({ filename: '02.mkv', season: 1, episode: 2 }); // the reload
    await flushAsyncWork();

    assert.equal(h.episodeInput.value, '2');
  });
});

test('a stale media-info reply from an earlier open does not make the current video look changed', async () => {
  await withReopenHarness(async (h) => {
    h.deferReplies();
    h.modal.openJimakuModal(); // load for episode 1 stays pending
    h.modal.closeJimakuModal();
    h.modal.openJimakuModal(); // load for episode 2
    h.pendingReplies[1]?.({ filename: '02.mkv', season: 1, episode: 2 });
    await flushAsyncWork();
    h.pendingReplies[0]?.({ filename: '01.mkv', season: 1, episode: 1 }); // stale
    await flushAsyncWork();
    h.episodeInput.value = '7'; // the user's own edit

    h.modal.openJimakuModal(); // same video: must keep the user's search
    h.pendingReplies[2]?.({ filename: '02.mkv', season: 1, episode: 2 });
    await flushAsyncWork();
    h.pendingReplies[3]?.({ filename: '02.mkv', season: 1, episode: 2 }); // a reload, if any
    await flushAsyncWork();

    assert.equal(h.pendingReplies.length, 3);
    assert.equal(h.episodeInput.value, '7');
  });
});

import assert from 'node:assert/strict';
import test from 'node:test';

import type { AnimetoshoSubtitleFile, ElectronAPI } from '../../types';
import { createRendererState } from '../state.js';
import { createAnimetoshoModal } from './animetosho.js';

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

const ENGLISH_TRACK: AnimetoshoSubtitleFile = {
  attachmentId: 1955356,
  filename: 'episode01.eng.ass',
  lang: 'eng',
  trackName: 'English subs',
  size: 33075,
  url: 'https://animetosho.org/storage/attach/001dd61c/1955356.xz',
  sourceFilename: 'episode01.mkv',
};

const JAPANESE_TRACK: AnimetoshoSubtitleFile = {
  attachmentId: 1955400,
  filename: 'episode01.jpn.ass',
  lang: 'jpn',
  trackName: 'Japanese subs',
  size: 41000,
  url: 'https://animetosho.org/storage/attach/001dd648/1955400.xz',
  sourceFilename: 'episode01.mkv',
};

const GERMAN_TRACK: AnimetoshoSubtitleFile = {
  attachmentId: 1955500,
  filename: 'episode01.ger.ass',
  lang: 'ger',
  trackName: 'Deutsch',
  size: 28000,
  url: 'https://animetosho.org/storage/attach/001dd6ac/1955500.xz',
  sourceFilename: 'episode01.mkv',
};

interface ModalHarness {
  modal: ReturnType<typeof createAnimetoshoModal>;
  state: ReturnType<typeof createRendererState>;
  downloadQueries: unknown[];
  modalCloseNotifications: string[];
  overlayClassList: ReturnType<typeof createClassList>;
  animetoshoModalClassList: ReturnType<typeof createClassList>;
  titleInput: { value: string };
  status: { textContent: string; style: { color: string } };
  entriesList: ReturnType<typeof createListStub>;
  filesList: ReturnType<typeof createListStub>;
  restoreGlobals: () => void;
}

function createModalHarness(
  files: AnimetoshoSubtitleFile[],
  options: {
    secondaryLanguages?: string[];
    listFiles?: (entryId: number) => Promise<unknown>;
    searchEntries?: (query: unknown) => Promise<unknown>;
  } = {},
): ModalHarness {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const hadWindow = Object.prototype.hasOwnProperty.call(globalThis, 'window');
  const hadDocument = Object.prototype.hasOwnProperty.call(globalThis, 'document');
  const previousWindow = globals.window;
  const previousDocument = globals.document;

  const modalCloseNotifications: string[] = [];
  const downloadQueries: unknown[] = [];

  const electronAPI = {
    animetoshoDownloadFile: async (query: unknown) => {
      downloadQueries.push(query);
      return { ok: true, path: '/tmp/subtitles/episode01.en.ass' };
    },
    animetoshoGetSecondaryLanguages: async () => options.secondaryLanguages ?? ['en', 'eng'],
    animetoshoListFiles: async ({ entryId }: { entryId: number }) =>
      options.listFiles ? options.listFiles(entryId) : { ok: true, data: [] },
    animetoshoSearchEntries: async (query: unknown) =>
      options.searchEntries ? options.searchEntries(query) : { ok: true, data: [] },
    getJimakuMediaInfo: async () => ({
      title: '',
      season: null,
      episode: null,
      confidence: 'low',
      filename: '',
      rawTitle: '',
    }),
    notifyOverlayModalClosed: (modal: string) => {
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

  const overlayClassList = createClassList(['interactive']);
  const animetoshoModalClassList = createClassList();
  const state = createRendererState();
  state.animetoshoModalOpen = true;
  state.currentAnimetoshoEntryId = 606713;
  state.selectedAnimetoshoFileIndex = 0;
  state.animetoshoFiles = files;

  const ctx = {
    dom: {
      overlay: { classList: overlayClassList },
      animetoshoModal: {
        classList: animetoshoModalClassList,
        setAttribute: () => {},
      },
      animetoshoTitleInput: { value: '' },
      animetoshoEpisodeInput: { value: '' },
      animetoshoSearchButton: { addEventListener: () => {} },
      animetoshoCloseButton: { addEventListener: () => {} },
      animetoshoTabEnglishButton: {
        textContent: 'English',
        classList: createClassList(['active']),
        addEventListener: () => {},
      },
      animetoshoTabJapaneseButton: {
        textContent: 'Japanese',
        classList: createClassList(),
        addEventListener: () => {},
      },
      animetoshoStatus: { textContent: '', style: { color: '' } },
      animetoshoEntriesSection: { classList: createClassList(['hidden']) },
      animetoshoEntriesList: createListStub(),
      animetoshoFilesSection: { classList: createClassList() },
      animetoshoFilesList: createListStub(),
    },
    state,
  };

  const modal = createAnimetoshoModal(ctx as never, {
    modalStateReader: { isAnyModalOpen: () => false },
    syncSettingsModalSubtitleSuppression: () => {},
  });

  return {
    modal,
    state,
    downloadQueries,
    modalCloseNotifications,
    overlayClassList,
    animetoshoModalClassList,
    titleInput: ctx.dom.animetoshoTitleInput,
    status: ctx.dom.animetoshoStatus,
    entriesList: ctx.dom.animetoshoEntriesList,
    filesList: ctx.dom.animetoshoFilesList,
    restoreGlobals: () => {
      const target = globalThis as unknown as Record<string, unknown>;
      if (hadWindow) {
        Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
      } else {
        delete target.window;
      }
      if (hadDocument) {
        Object.defineProperty(globalThis, 'document', {
          configurable: true,
          value: previousDocument,
        });
      } else {
        delete target.document;
      }
    },
  };
}

function pressKey(harness: ModalHarness, key: string): boolean {
  let prevented = false;
  harness.modal.handleAnimetoshoKeydown({
    key,
    preventDefault: () => {
      prevented = true;
    },
  } as KeyboardEvent);
  return prevented;
}

test('successful Animetosho subtitle selection closes modal', async () => {
  const harness = createModalHarness([ENGLISH_TRACK, JAPANESE_TRACK]);
  try {
    const prevented = pressKey(harness, 'Enter');
    await flushAsyncWork();

    assert.equal(prevented, true);
    assert.equal(harness.state.animetoshoModalOpen, false);
    assert.equal(harness.animetoshoModalClassList.contains('hidden'), true);
    assert.equal(harness.overlayClassList.contains('interactive'), false);
    assert.deepEqual(harness.modalCloseNotifications, ['animetosho']);
    assert.deepEqual(harness.downloadQueries, [
      {
        entryId: 606713,
        url: ENGLISH_TRACK.url,
        name: ENGLISH_TRACK.filename,
        lang: 'eng',
      },
    ]);
  } finally {
    harness.restoreGlobals();
  }
});

test('English tab hides non-English languages, not just Japanese', async () => {
  const harness = createModalHarness([GERMAN_TRACK, ENGLISH_TRACK, JAPANESE_TRACK]);
  try {
    // With German visible this would move selection onto it; English-only
    // filtering must clamp to the single English track instead.
    pressKey(harness, 'ArrowDown');
    pressKey(harness, 'Enter');
    await flushAsyncWork();

    assert.deepEqual(harness.downloadQueries, [
      {
        entryId: 606713,
        url: ENGLISH_TRACK.url,
        name: ENGLISH_TRACK.filename,
        lang: 'eng',
      },
    ]);
  } finally {
    harness.restoreGlobals();
  }
});

test('Japanese tab filters tracks so Enter downloads the Japanese one', async () => {
  const harness = createModalHarness([ENGLISH_TRACK, JAPANESE_TRACK]);
  try {
    assert.equal(harness.state.animetoshoActiveTab, 'en');
    pressKey(harness, 'ArrowRight');
    assert.equal(harness.state.animetoshoActiveTab, 'ja');

    pressKey(harness, 'Enter');
    await flushAsyncWork();

    assert.deepEqual(harness.downloadQueries, [
      {
        entryId: 606713,
        url: JAPANESE_TRACK.url,
        name: JAPANESE_TRACK.filename,
        lang: 'jpn',
      },
    ]);
  } finally {
    harness.restoreGlobals();
  }
});

test('secondary tab follows configured secondarySub languages', async () => {
  const harness = createModalHarness([GERMAN_TRACK, ENGLISH_TRACK, JAPANESE_TRACK], {
    secondaryLanguages: ['de'],
  });
  try {
    // Re-open through the API so the modal fetches the configured languages.
    harness.state.animetoshoModalOpen = false;
    harness.modal.openAnimetoshoModal();
    await flushAsyncWork();

    harness.state.animetoshoFiles = [GERMAN_TRACK, ENGLISH_TRACK, JAPANESE_TRACK];
    harness.state.currentAnimetoshoEntryId = 606713;

    pressKey(harness, 'Enter');
    await flushAsyncWork();

    assert.deepEqual(harness.downloadQueries, [
      {
        entryId: 606713,
        url: GERMAN_TRACK.url,
        name: GERMAN_TRACK.filename,
        lang: 'ger',
      },
    ]);
  } finally {
    harness.restoreGlobals();
  }
});

test('a slow release response does not overwrite the newly selected release', async () => {
  const STALE_TRACK: AnimetoshoSubtitleFile = {
    ...ENGLISH_TRACK,
    attachmentId: 999,
    filename: 'stale.eng.ass',
  };
  const SECOND_ENGLISH_TRACK: AnimetoshoSubtitleFile = {
    ...ENGLISH_TRACK,
    attachmentId: 1955357,
    filename: 'episode01.eng.sdh.ass',
  };
  const resolvers: Array<(value: unknown) => void> = [];

  const harness = createModalHarness([], {
    listFiles: (entryId) =>
      new Promise((resolve) => {
        if (entryId === 1) {
          // Entry 1 answers late, after the user has moved on to entry 2.
          resolvers.push(() => resolve({ ok: true, data: [STALE_TRACK] }));
        } else {
          // Two tracks, so the modal does not auto-download a lone match.
          resolve({ ok: true, data: [ENGLISH_TRACK, SECOND_ENGLISH_TRACK] });
        }
      }),
  });

  try {
    harness.state.animetoshoEntries = [
      { id: 1, title: 'slow release', timestamp: null, totalSize: null, numFiles: 1, sublangs: [] },
      { id: 2, title: 'fast release', timestamp: null, totalSize: null, numFiles: 1, sublangs: [] },
    ];

    harness.modal.selectAnimetoshoEntry(0);
    harness.modal.selectAnimetoshoEntry(1);
    await flushAsyncWork();

    // Entry 2's tracks are on screen; now entry 1 finally answers.
    assert.deepEqual(
      harness.state.animetoshoFiles.map((file) => file.attachmentId),
      [ENGLISH_TRACK.attachmentId, SECOND_ENGLISH_TRACK.attachmentId],
    );

    resolvers.forEach((resolve) => resolve(undefined));
    await flushAsyncWork();

    assert.equal(harness.state.currentAnimetoshoEntryId, 2);
    assert.deepEqual(
      harness.state.animetoshoFiles.map((file) => file.attachmentId),
      [ENGLISH_TRACK.attachmentId, SECOND_ENGLISH_TRACK.attachmentId],
    );
  } finally {
    harness.restoreGlobals();
  }
});

test('ArrowLeft switches back to the English tab', () => {
  const harness = createModalHarness([ENGLISH_TRACK, JAPANESE_TRACK]);
  try {
    pressKey(harness, 'ArrowRight');
    assert.equal(harness.state.animetoshoActiveTab, 'ja');
    pressKey(harness, 'ArrowLeft');
    assert.equal(harness.state.animetoshoActiveTab, 'en');
  } finally {
    harness.restoreGlobals();
  }
});

test('searching reports TsukiHime as the backend and lists sublangs per release', async () => {
  const harness = createModalHarness([], {
    searchEntries: async () => ({
      ok: true,
      data: [
        {
          id: 12255,
          title: '[DKB] Futsutsuka na Akujo - S01E01 [Multi-Subs]',
          timestamp: null,
          totalSize: 423868898,
          numFiles: 1,
          sublangs: ['en-US', 'ja'],
        },
        {
          id: 12256,
          title: 'release without langs',
          timestamp: null,
          totalSize: null,
          numFiles: null,
          sublangs: [],
        },
      ],
    }),
  });
  try {
    harness.state.animetoshoFiles = [];
    harness.state.currentAnimetoshoEntryId = null;
    harness.titleInput.value = 'Futsutsuka na Akujo';

    const seenStatuses: string[] = [];
    const status = harness.status;
    Object.defineProperty(status, 'textContent', {
      get: () => seenStatuses.at(-1) ?? '',
      set: (value: string) => {
        seenStatuses.push(value);
      },
    });

    pressKey(harness, 'Enter');
    await flushAsyncWork();

    assert.equal(
      seenStatuses.some((message) => message.includes('TsukiHime')),
      true,
    );
    assert.equal(
      seenStatuses.some((message) => message.includes('Animetosho')),
      false,
    );

    const firstEntry = harness.entriesList.children[0] as {
      children: Array<{ textContent: string }>;
    };
    assert.equal(firstEntry.children.length, 1);
    assert.match(firstEntry.children[0]!.textContent, /en-US, ja/);
  } finally {
    harness.restoreGlobals();
  }
});

test('renderFiles omits the size detail when the API does not report one', () => {
  const zeroSizeTrack: AnimetoshoSubtitleFile = {
    ...ENGLISH_TRACK,
    size: 0,
  };
  const harness = createModalHarness([zeroSizeTrack, JAPANESE_TRACK]);
  try {
    pressKey(harness, 'ArrowDown');

    const firstFile = harness.filesList.children[0] as {
      children: Array<{ textContent: string }>;
    };
    assert.equal(firstFile.children.length, 1);
    assert.equal(firstFile.children[0]!.textContent.includes('0 B'), false);
    assert.match(firstFile.children[0]!.textContent, /English subs/);
  } finally {
    harness.restoreGlobals();
  }
});

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
  restoreGlobals: () => void;
}

function createModalHarness(
  files: AnimetoshoSubtitleFile[],
  options: { secondaryLanguages?: string[] } = {},
): ModalHarness {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
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
    restoreGlobals: () => {
      Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
      Object.defineProperty(globalThis, 'document', {
        configurable: true,
        value: previousDocument,
      });
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

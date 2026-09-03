import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';

import type { TsukihimeSubtitleFile, ElectronAPI } from '../../types';
import { createRendererState } from '../state.js';
import { createTsukihimeModal } from './tsukihime.js';

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
  const list = {
    children: [] as unknown[],
    appendChild(child: unknown) {
      list.children.push(child);
    },
  };
  // The modal clears lists through innerHTML before re-rendering.
  return Object.defineProperty(list, 'innerHTML', {
    get: () => '',
    set: () => {
      list.children.length = 0;
    },
  }) as typeof list & { innerHTML: string };
}

function createTabStub(active: boolean) {
  const attributes = new Map<string, string>();
  return {
    textContent: '',
    classList: createClassList(active ? ['active'] : []),
    attributes,
    setAttribute(name: string, value: string) {
      attributes.set(name, value);
    },
    addEventListener: () => {},
  };
}

function flushAsyncWork(): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, 0);
  });
}

const ENGLISH_TRACK: TsukihimeSubtitleFile = {
  attachmentId: 1955356,
  filename: 'episode01.eng.ass',
  lang: 'eng',
  trackName: 'English subs',
  size: 33075,
  url: 'https://storage.tsukihime.org/attach/001dd61c/1955356.xz',
  sourceFilename: 'episode01.mkv',
};

const JAPANESE_TRACK: TsukihimeSubtitleFile = {
  attachmentId: 1955400,
  filename: 'episode01.jpn.ass',
  lang: 'jpn',
  trackName: 'Japanese subs',
  size: 41000,
  url: 'https://storage.tsukihime.org/attach/001dd648/1955400.xz',
  sourceFilename: 'episode01.mkv',
};

const GERMAN_TRACK: TsukihimeSubtitleFile = {
  attachmentId: 1955500,
  filename: 'episode01.ger.ass',
  lang: 'ger',
  trackName: 'Deutsch',
  size: 28000,
  url: 'https://storage.tsukihime.org/attach/001dd6ac/1955500.xz',
  sourceFilename: 'episode01.mkv',
};

interface ModalHarness {
  modal: ReturnType<typeof createTsukihimeModal>;
  state: ReturnType<typeof createRendererState>;
  downloadQueries: unknown[];
  modalCloseNotifications: string[];
  overlayClassList: ReturnType<typeof createClassList>;
  tsukihimeModalClassList: ReturnType<typeof createClassList>;
  titleInput: { value: string };
  status: { textContent: string; style: { color: string } };
  entriesList: ReturnType<typeof createListStub>;
  filesList: ReturnType<typeof createListStub>;
  secondaryTab: ReturnType<typeof createTabStub>;
  primaryTab: ReturnType<typeof createTabStub>;
  restoreGlobals: () => void;
}

function createModalHarness(
  files: TsukihimeSubtitleFile[],
  options: {
    secondaryLanguages?: string[];
    secondaryLanguagesGate?: Promise<void>;
    downloadFile?: (query: unknown) => Promise<unknown>;
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
    tsukihimeDownloadFile: async (query: unknown) => {
      downloadQueries.push(query);
      if (options.downloadFile) return options.downloadFile(query);
      return { ok: true, path: '/tmp/subtitles/episode01.en.ass' };
    },
    tsukihimeGetSecondaryLanguages: async () => {
      await options.secondaryLanguagesGate;
      return options.secondaryLanguages ?? ['en', 'eng'];
    },
    tsukihimeListFiles: async ({ entryId }: { entryId: number }) =>
      options.listFiles ? options.listFiles(entryId) : { ok: true, data: [] },
    tsukihimeSearchEntries: async (query: unknown) =>
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
  const tsukihimeModalClassList = createClassList();
  const state = createRendererState();
  state.tsukihimeModalOpen = true;
  state.currentTsukihimeEntryId = 606713;
  state.selectedTsukihimeFileIndex = 0;
  state.tsukihimeFiles = files;
  const secondaryTab = createTabStub(true);
  secondaryTab.textContent = 'English';
  const primaryTab = createTabStub(false);
  primaryTab.textContent = 'Japanese';

  const ctx = {
    dom: {
      overlay: { classList: overlayClassList },
      tsukihimeModal: {
        classList: tsukihimeModalClassList,
        setAttribute: () => {},
      },
      tsukihimeTitleInput: { value: '' },
      tsukihimeSeasonInput: { value: '' },
      tsukihimeEpisodeInput: { value: '' },
      tsukihimeSearchButton: { addEventListener: () => {} },
      tsukihimeCloseButton: { addEventListener: () => {} },
      tsukihimeTabSecondaryButton: secondaryTab,
      tsukihimeTabPrimaryButton: primaryTab,
      tsukihimeStatus: { textContent: '', style: { color: '' } },
      tsukihimeEntriesSection: { classList: createClassList(['hidden']) },
      tsukihimeEntriesList: createListStub(),
      tsukihimeFilesSection: { classList: createClassList() },
      tsukihimeFilesList: createListStub(),
    },
    state,
  };

  const modal = createTsukihimeModal(ctx as never, {
    modalStateReader: { isAnyModalOpen: () => false },
    syncSettingsModalSubtitleSuppression: () => {},
  });

  return {
    modal,
    state,
    downloadQueries,
    modalCloseNotifications,
    overlayClassList,
    tsukihimeModalClassList,
    titleInput: ctx.dom.tsukihimeTitleInput,
    status: ctx.dom.tsukihimeStatus,
    entriesList: ctx.dom.tsukihimeEntriesList,
    filesList: ctx.dom.tsukihimeFilesList,
    secondaryTab,
    primaryTab,
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

test('TsukiHime language tabs expose tab semantics in the renderer markup', () => {
  const html = fs.readFileSync(path.join(process.cwd(), 'src', 'renderer', 'index.html'), 'utf8');
  const tabs = html.match(/<div class="tsukihime-tabs"[\s\S]*?<\/div>/)?.[0];

  assert.ok(tabs);
  assert.match(tabs, /<div class="tsukihime-tabs" role="tablist">/);
  assert.match(tabs, /id="tsukihimeTabSecondary"[\s\S]*?role="tab"[\s\S]*?aria-selected="true"/);
  assert.match(tabs, /id="tsukihimeTabPrimary"[\s\S]*?role="tab"[\s\S]*?aria-selected="false"/);
});

test('switching TsukiHime language tabs synchronizes aria-selected', () => {
  const harness = createModalHarness([ENGLISH_TRACK, JAPANESE_TRACK]);
  try {
    pressKey(harness, 'ArrowRight');

    assert.equal(harness.secondaryTab.attributes.get('aria-selected'), 'false');
    assert.equal(harness.primaryTab.attributes.get('aria-selected'), 'true');

    pressKey(harness, 'ArrowLeft');

    assert.equal(harness.secondaryTab.attributes.get('aria-selected'), 'true');
    assert.equal(harness.primaryTab.attributes.get('aria-selected'), 'false');
  } finally {
    harness.restoreGlobals();
  }
});

function pressKey(harness: ModalHarness, key: string): boolean {
  let prevented = false;
  harness.modal.handleTsukihimeKeydown({
    key,
    preventDefault: () => {
      prevented = true;
    },
  } as KeyboardEvent);
  return prevented;
}

test('successful Tsukihime subtitle selection closes modal', async () => {
  const harness = createModalHarness([ENGLISH_TRACK, JAPANESE_TRACK]);
  try {
    const prevented = pressKey(harness, 'Enter');
    await flushAsyncWork();

    assert.equal(prevented, true);
    assert.equal(harness.state.tsukihimeModalOpen, false);
    assert.equal(harness.tsukihimeModalClassList.contains('hidden'), true);
    assert.equal(harness.overlayClassList.contains('interactive'), false);
    assert.deepEqual(harness.modalCloseNotifications, ['tsukihime']);
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

test('a download from a prior modal session cannot close a reopened modal', async () => {
  let resolveDownload!: (value: unknown) => void;
  const harness = createModalHarness([ENGLISH_TRACK, JAPANESE_TRACK], {
    downloadFile: () =>
      new Promise((resolve) => {
        resolveDownload = resolve;
      }),
  });
  try {
    pressKey(harness, 'Enter');
    harness.modal.closeTsukihimeModal();
    harness.modal.openTsukihimeModal();
    await flushAsyncWork();
    harness.state.currentTsukihimeEntryId = 606713;
    harness.status.textContent = 'Fresh modal session';

    resolveDownload({ ok: true, path: '/tmp/subtitles/stale.ass' });
    await flushAsyncWork();

    assert.equal(harness.state.tsukihimeModalOpen, true);
    assert.equal(harness.status.textContent, 'Fresh modal session');
    assert.deepEqual(harness.modalCloseNotifications, ['tsukihime']);
  } finally {
    harness.restoreGlobals();
  }
});

test('a download cannot affect a newly selected release', async () => {
  let resolveDownload!: (value: unknown) => void;
  const harness = createModalHarness([ENGLISH_TRACK, JAPANESE_TRACK], {
    downloadFile: () =>
      new Promise((resolve) => {
        resolveDownload = resolve;
      }),
  });
  try {
    pressKey(harness, 'Enter');
    harness.state.tsukihimeEntries = [
      {
        id: 999,
        title: 'new release',
        timestamp: null,
        totalSize: null,
        numFiles: 1,
        sublangs: [],
      },
    ];
    harness.modal.selectTsukihimeEntry(0);
    await flushAsyncWork();
    const currentStatus = harness.status.textContent;

    resolveDownload({ ok: false, error: { error: 'stale failure' } });
    await flushAsyncWork();

    assert.equal(harness.state.tsukihimeModalOpen, true);
    assert.equal(harness.state.currentTsukihimeEntryId, 999);
    assert.equal(harness.status.textContent, currentStatus);
    assert.deepEqual(harness.modalCloseNotifications, []);
  } finally {
    harness.restoreGlobals();
  }
});

test('secondary tab hides languages outside its configured defaults', async () => {
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

test('primary tab defaults to Japanese tracks', async () => {
  const harness = createModalHarness([ENGLISH_TRACK, JAPANESE_TRACK]);
  try {
    assert.equal(harness.state.tsukihimeActiveTab, 'secondary');
    pressKey(harness, 'ArrowRight');
    assert.equal(harness.state.tsukihimeActiveTab, 'primary');

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
    harness.state.tsukihimeModalOpen = false;
    harness.modal.openTsukihimeModal();
    await flushAsyncWork();

    harness.state.tsukihimeFiles = [GERMAN_TRACK, ENGLISH_TRACK, JAPANESE_TRACK];
    harness.state.currentTsukihimeEntryId = 606713;

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

test('primary tab remains Japanese when the release includes other languages', async () => {
  const harness = createModalHarness([GERMAN_TRACK, ENGLISH_TRACK, JAPANESE_TRACK], {
    secondaryLanguages: ['en'],
  });
  try {
    harness.state.tsukihimeModalOpen = false;
    harness.modal.openTsukihimeModal();
    await flushAsyncWork();

    harness.state.tsukihimeFiles = [GERMAN_TRACK, ENGLISH_TRACK, JAPANESE_TRACK];
    harness.state.currentTsukihimeEntryId = 606713;

    pressKey(harness, 'ArrowRight');
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

test('a slow release response does not overwrite the newly selected release', async () => {
  const STALE_TRACK: TsukihimeSubtitleFile = {
    ...ENGLISH_TRACK,
    attachmentId: 999,
    filename: 'stale.eng.ass',
  };
  const SECOND_ENGLISH_TRACK: TsukihimeSubtitleFile = {
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
    harness.state.tsukihimeEntries = [
      { id: 1, title: 'slow release', timestamp: null, totalSize: null, numFiles: 1, sublangs: [] },
      { id: 2, title: 'fast release', timestamp: null, totalSize: null, numFiles: 1, sublangs: [] },
    ];

    harness.modal.selectTsukihimeEntry(0);
    harness.modal.selectTsukihimeEntry(1);
    await flushAsyncWork();

    // Entry 2's tracks are on screen; now entry 1 finally answers.
    assert.deepEqual(
      harness.state.tsukihimeFiles.map((file) => file.attachmentId),
      [ENGLISH_TRACK.attachmentId, SECOND_ENGLISH_TRACK.attachmentId],
    );

    resolvers.forEach((resolve) => resolve(undefined));
    await flushAsyncWork();

    assert.equal(harness.state.currentTsukihimeEntryId, 2);
    assert.deepEqual(
      harness.state.tsukihimeFiles.map((file) => file.attachmentId),
      [ENGLISH_TRACK.attachmentId, SECOND_ENGLISH_TRACK.attachmentId],
    );
  } finally {
    harness.restoreGlobals();
  }
});

test('ArrowLeft switches back to the secondary-language tab', () => {
  const harness = createModalHarness([ENGLISH_TRACK, JAPANESE_TRACK]);
  try {
    pressKey(harness, 'ArrowRight');
    assert.equal(harness.state.tsukihimeActiveTab, 'primary');
    pressKey(harness, 'ArrowLeft');
    assert.equal(harness.state.tsukihimeActiveTab, 'secondary');
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
    harness.state.tsukihimeFiles = [];
    harness.state.currentTsukihimeEntryId = null;
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
      seenStatuses.some((message) => message.includes('Tsukihime')),
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
  const zeroSizeTrack: TsukihimeSubtitleFile = {
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

const ENGLISH_ONLY_ENTRY = {
  id: 606713,
  title: 'english only release',
  timestamp: null,
  totalSize: null,
  numFiles: 1,
  sublangs: ['en'],
};

const MULTI_SUB_ENTRY = {
  id: 12255,
  title: 'multi-sub release',
  timestamp: null,
  totalSize: null,
  numFiles: 1,
  sublangs: ['en-US', 'ja'],
};

const UNLABELED_ENTRY = {
  id: 12256,
  title: 'release without langs',
  timestamp: null,
  totalSize: null,
  numFiles: 1,
  sublangs: [],
};

function visibleEntryTitles(harness: ModalHarness): string[] {
  return (harness.entriesList.children as Array<{ textContent: string }>).map(
    (li) => li.textContent,
  );
}

test('Japanese tab lists only releases that carry Japanese subtitles', async () => {
  const SECOND_JAPANESE_TRACK: TsukihimeSubtitleFile = {
    ...JAPANESE_TRACK,
    attachmentId: 1955401,
    filename: 'episode01.jpn.sdh.ass',
  };
  const harness = createModalHarness([], {
    // Two tracks so the modal does not auto-download a lone match.
    listFiles: async () => ({ ok: true, data: [JAPANESE_TRACK, SECOND_JAPANESE_TRACK] }),
  });
  try {
    harness.state.currentTsukihimeEntryId = null;
    harness.state.tsukihimeEntries = [ENGLISH_ONLY_ENTRY, MULTI_SUB_ENTRY, UNLABELED_ENTRY];

    pressKey(harness, 'ArrowRight');
    assert.deepEqual(visibleEntryTitles(harness), ['multi-sub release']);

    // Enter addresses the visible list, so it must pick the multi-sub release
    // rather than the hidden first search result.
    pressKey(harness, 'Enter');
    await flushAsyncWork();
    assert.equal(harness.state.currentTsukihimeEntryId, MULTI_SUB_ENTRY.id);
    assert.equal(harness.status.textContent, 'Select a subtitle track.');

    pressKey(harness, 'ArrowLeft');
    assert.deepEqual(visibleEntryTitles(harness), [
      'english only release',
      'multi-sub release',
      'release without langs',
    ]);
    assert.equal(harness.state.currentTsukihimeEntryId, MULTI_SUB_ENTRY.id);
    assert.equal(harness.state.selectedTsukihimeEntryIndex, 1);
  } finally {
    harness.restoreGlobals();
  }
});

test('Japanese tab reports when no release carries Japanese subtitles', () => {
  const harness = createModalHarness([]);
  try {
    harness.state.currentTsukihimeEntryId = null;
    harness.state.tsukihimeEntries = [ENGLISH_ONLY_ENTRY, UNLABELED_ENTRY];

    pressKey(harness, 'ArrowRight');
    assert.deepEqual(visibleEntryTitles(harness), []);
    assert.equal(
      harness.status.textContent,
      'No releases with Japanese subtitles. Switch to the English tab.',
    );

    pressKey(harness, 'ArrowLeft');
    assert.deepEqual(visibleEntryTitles(harness), [
      'english only release',
      'release without langs',
    ]);
    assert.equal(harness.status.textContent, 'Select a release.');
  } finally {
    harness.restoreGlobals();
  }
});

test('search reports when no release carries the secondary language', async () => {
  const harness = createModalHarness([], {
    searchEntries: async () => ({
      ok: true,
      data: [{ ...MULTI_SUB_ENTRY, sublangs: ['ja'] }],
    }),
  });
  try {
    harness.state.currentTsukihimeEntryId = null;
    harness.titleInput.value = 'Futsutsuka na Akujo';

    pressKey(harness, 'Enter');
    await flushAsyncWork();
    assert.deepEqual(visibleEntryTitles(harness), []);
    assert.equal(
      harness.status.textContent,
      'No releases with English subtitles. Switch to the Japanese tab.',
    );

    pressKey(harness, 'ArrowRight');
    assert.deepEqual(visibleEntryTitles(harness), ['multi-sub release']);
  } finally {
    harness.restoreGlobals();
  }
});

test('switching to a tab that hides the selected release clears its tracks', () => {
  const harness = createModalHarness([ENGLISH_TRACK, JAPANESE_TRACK]);
  try {
    harness.state.tsukihimeEntries = [ENGLISH_ONLY_ENTRY, MULTI_SUB_ENTRY];

    pressKey(harness, 'ArrowRight');
    assert.equal(harness.state.currentTsukihimeEntryId, null);
    assert.deepEqual(harness.state.tsukihimeFiles, []);
    assert.deepEqual(visibleEntryTitles(harness), ['multi-sub release']);
    assert.equal(harness.status.textContent, 'Select a release.');
  } finally {
    harness.restoreGlobals();
  }
});

test('a search waits for the configured secondary languages before filtering', async () => {
  let openGate!: () => void;
  const harness = createModalHarness([], {
    secondaryLanguages: ['de'],
    secondaryLanguagesGate: new Promise<void>((resolve) => {
      openGate = resolve;
    }),
    searchEntries: async () => ({
      ok: true,
      data: [{ ...MULTI_SUB_ENTRY, title: 'german release', sublangs: ['de'] }],
    }),
  });
  try {
    harness.state.tsukihimeModalOpen = false;
    harness.modal.openTsukihimeModal();
    harness.titleInput.value = 'Futsutsuka na Akujo';

    // Searching before the config arrives must not filter against the English
    // fallback, which would hide this German-only release.
    pressKey(harness, 'Enter');
    await flushAsyncWork();
    assert.deepEqual(visibleEntryTitles(harness), []);

    openGate();
    await flushAsyncWork();

    assert.deepEqual(visibleEntryTitles(harness), ['german release']);
  } finally {
    harness.restoreGlobals();
  }
});

test('a search from a prior modal session cannot repopulate a reopened modal', async () => {
  let openGate!: () => void;
  const harness = createModalHarness([], {
    secondaryLanguagesGate: new Promise<void>((resolve) => {
      openGate = resolve;
    }),
    searchEntries: async () => ({ ok: true, data: [MULTI_SUB_ENTRY] }),
  });
  try {
    harness.state.tsukihimeModalOpen = false;
    harness.modal.openTsukihimeModal();
    harness.titleInput.value = 'Futsutsuka na Akujo';

    // The search parks on the language config, then the user closes and
    // reopens the modal before it resolves.
    pressKey(harness, 'Enter');
    harness.modal.closeTsukihimeModal();
    harness.modal.openTsukihimeModal();
    await flushAsyncWork();
    harness.status.textContent = 'Fresh modal session';

    openGate();
    await flushAsyncWork();

    assert.deepEqual(harness.state.tsukihimeEntries, []);
    assert.deepEqual(visibleEntryTitles(harness), []);
    assert.equal(harness.status.textContent, 'Fresh modal session');
  } finally {
    harness.restoreGlobals();
  }
});

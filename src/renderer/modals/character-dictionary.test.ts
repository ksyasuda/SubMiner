import assert from 'node:assert/strict';
import test from 'node:test';

import type { CharacterDictionarySelectionSnapshot, ElectronAPI } from '../../types';
import { createRendererState } from '../state.js';
import { createCharacterDictionaryModal } from './character-dictionary.js';

function createClassList(initialTokens: string[] = []) {
  const tokens = new Set(initialTokens);
  return {
    add: (...entries: string[]) => entries.forEach((entry) => tokens.add(entry)),
    remove: (...entries: string[]) => entries.forEach((entry) => tokens.delete(entry)),
    toggle: (entry: string, force?: boolean) => {
      if (force === undefined) {
        if (tokens.has(entry)) tokens.delete(entry);
        else tokens.add(entry);
        return;
      }
      if (force) tokens.add(entry);
      else tokens.delete(entry);
    },
    contains: (entry: string) => tokens.has(entry),
  };
}

type EventPayload = { stopPropagation?: () => void; preventDefault?: () => void };

// Covers both `ctx.dom` nodes and elements the modal builds via `document.createElement`.
function createNodeStub(hidden = false) {
  const listeners = new Map<string, Array<(event?: EventPayload) => void>>();
  return {
    className: '',
    textContent: '',
    type: '',
    value: '',
    disabled: false,
    children: [] as unknown[],
    classList: createClassList(hidden ? ['hidden'] : []),
    setAttribute: () => {},
    addEventListener: (event: string, listener: (event?: EventPayload) => void) => {
      listeners.set(event, [...(listeners.get(event) ?? []), listener]);
    },
    dispatchEvent: (event: string, payload?: EventPayload) => {
      for (const listener of listeners.get(event) ?? []) listener(payload);
    },
    append(...children: unknown[]) {
      this.children.push(...children);
    },
    replaceChildren(...children: unknown[]) {
      this.children = [...children];
    },
  };
}

type NodeStub = ReturnType<typeof createNodeStub>;

type CharacterDictionaryApi = Pick<
  ElectronAPI,
  | 'getCharacterDictionarySelection'
  | 'setCharacterDictionarySelection'
  | 'getCharacterDictionaryManagerSnapshot'
  | 'removeCharacterDictionaryManagedEntry'
  | 'moveCharacterDictionaryManagedEntry'
  | 'notifyOverlayModalClosed'
  | 'notifyOverlayModalOpened'
>;

function createDom() {
  return {
    overlay: createNodeStub(),
    characterDictionaryModal: createNodeStub(true),
    characterDictionaryClose: createNodeStub(),
    characterDictionarySummary: createNodeStub(),
    characterDictionaryCurrent: createNodeStub(),
    characterDictionarySearchInput: createNodeStub(),
    characterDictionarySearchButton: createNodeStub(),
    characterDictionaryCandidates: createNodeStub(),
    characterDictionaryStatus: createNodeStub(),
    characterDictionarySearchPanel: createNodeStub(),
    characterDictionaryManagerPanel: createNodeStub(true),
    characterDictionaryOverrideTab: createNodeStub(),
    characterDictionaryManageTab: createNodeStub(),
    characterDictionaryManagedEntries: createNodeStub(),
  };
}

type Harness = {
  modal: ReturnType<typeof createCharacterDictionaryModal>;
  state: ReturnType<typeof createRendererState>;
  dom: ReturnType<typeof createDom>;
  openedModals: string[];
  suppressionSyncs: () => number;
};

// Installs fake window/document globals, builds the modal, runs `body`, then restores globals.
async function withHarness(
  electronAPI: Partial<CharacterDictionaryApi>,
  body: (harness: Harness) => Promise<void>,
): Promise<void> {
  const previousWindow = globalThis.window;
  const previousDocument = globalThis.document;
  const openedModals: string[] = [];
  let suppressionSyncs = 0;
  const api: CharacterDictionaryApi = {
    getCharacterDictionarySelection: async () => ({
      seriesKey: '',
      guessTitle: null,
      current: null,
      override: null,
      candidates: [],
    }),
    setCharacterDictionarySelection: async () => ({
      ok: false,
      seriesKey: '',
      selected: { id: 0, title: '', episodes: null },
      staleMediaIds: [],
    }),
    getCharacterDictionaryManagerSnapshot: async () => ({ entries: [] }),
    removeCharacterDictionaryManagedEntry: async () => ({ ok: true, entries: [] }),
    moveCharacterDictionaryManagedEntry: async () => ({ ok: true, entries: [] }),
    notifyOverlayModalClosed: () => {},
    notifyOverlayModalOpened: (modal) => {
      openedModals.push(modal);
    },
    ...electronAPI,
  };
  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: { electronAPI: api },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: { createElement: () => createNodeStub() },
  });

  const state = createRendererState();
  const dom = createDom();
  const modal = createCharacterDictionaryModal({ state, dom } as never, {
    modalStateReader: { isAnyModalOpen: () => false },
    syncSettingsModalSubtitleSuppression: () => {
      suppressionSyncs += 1;
    },
  });

  try {
    await body({ modal, state, dom, openedModals, suppressionSyncs: () => suppressionSyncs });
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
}

const MANAGED_ENTRIES = [
  { mediaId: 21202, label: '21202 - KonoSuba', title: 'KonoSuba', current: true },
  { mediaId: 115230, label: '115230 - Tower of God', title: 'Tower of God', current: false },
];

// Manager entry controls are rendered as [Up, Down, Override, Remove].
function clickManagedEntryControl(
  managedEntries: NodeStub,
  entryIndex: number,
  controlIndex: number,
) {
  const entry = managedEntries.children[entryIndex] as NodeStub;
  const controls = entry.children[1] as NodeStub;
  (controls.children[controlIndex] as NodeStub).dispatchEvent('click', {
    stopPropagation: () => {},
  });
}

function pressEnter(modal: Harness['modal']): void {
  modal.handleCharacterDictionaryKeydown({
    key: 'Enter',
    preventDefault: () => {},
  } as KeyboardEvent);
}

function flushAsyncWork(): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, 0);
  });
}

test('character dictionary modal announces open before AniList refresh resolves', async () => {
  let resolveSelection: (snapshot: CharacterDictionarySelectionSnapshot) => void = () => {};
  const selectionPromise = new Promise<CharacterDictionarySelectionSnapshot>((resolve) => {
    resolveSelection = resolve;
  });

  await withHarness(
    { getCharacterDictionarySelection: () => selectionPromise },
    async ({ modal, state, dom, openedModals, suppressionSyncs }) => {
      const openPromise = modal.openCharacterDictionaryModal();

      assert.equal(state.characterDictionaryModalOpen, true);
      assert.equal(dom.characterDictionaryModal.classList.contains('hidden'), false);
      assert.equal(suppressionSyncs(), 1);
      assert.deepEqual(openedModals, ['character-dictionary']);

      resolveSelection({
        seriesKey: 'tower-of-god-2020',
        guessTitle: 'Tower of God',
        current: null,
        override: null,
        candidates: [{ id: 115230, title: 'Tower of God', episodes: 13 }],
      });
      await openPromise;
    },
  );
});

test('character dictionary modal opens manager view with active entries', async () => {
  await withHarness(
    { getCharacterDictionaryManagerSnapshot: async () => ({ entries: MANAGED_ENTRIES }) },
    async ({ modal, state, dom }) => {
      await modal.openCharacterDictionaryManagerModal();

      assert.equal(state.characterDictionaryModalOpen, true);
      assert.equal(dom.characterDictionaryManagedEntries.children.length, 2);
      assert.equal(
        dom.characterDictionarySummary.textContent,
        '2 loaded character dictionaries. Order controls eviction priority; current dictionary stays loaded.',
      );
    },
  );
});

test('character dictionary manager reports failed reorder IPC calls', async () => {
  await withHarness(
    {
      getCharacterDictionaryManagerSnapshot: async () => ({ entries: MANAGED_ENTRIES }),
      moveCharacterDictionaryManagedEntry: async () => {
        throw new Error('move failed');
      },
    },
    async ({ modal, dom }) => {
      await modal.openCharacterDictionaryManagerModal();
      clickManagedEntryControl(dom.characterDictionaryManagedEntries, 1, 0);
      await flushAsyncWork();

      assert.equal(dom.characterDictionaryStatus.textContent, 'move failed');
    },
  );
});

test('character dictionary manager reports pending refresh after removal', async () => {
  await withHarness(
    {
      getCharacterDictionaryManagerSnapshot: async () => ({ entries: MANAGED_ENTRIES }),
      removeCharacterDictionaryManagedEntry: async () => ({
        ok: true,
        entries: [MANAGED_ENTRIES[0]!],
        rebuildRequired: true,
      }),
    },
    async ({ modal, dom }) => {
      await modal.openCharacterDictionaryManagerModal();
      clickManagedEntryControl(dom.characterDictionaryManagedEntries, 1, 3);
      await flushAsyncWork();

      assert.equal(
        dom.characterDictionaryStatus.textContent,
        'Entry removed. Merged dictionary will refresh shortly.',
      );
    },
  );
});

test('character dictionary modal loads candidates and applies selected override', async () => {
  const snapshot: CharacterDictionarySelectionSnapshot = {
    seriesKey: 're-zero-starting-life-in-another-world-2016',
    guessTitle: 'Re ZERO, Starting Life in Another World',
    current: { id: 10607, title: 'Rerere no Tensai Bakabon', episodes: 24 },
    override: null,
    candidates: [{ id: 21355, title: 'Re:ZERO -Starting Life in Another World-', episodes: 25 }],
  };
  const savedIds: number[] = [];

  await withHarness(
    {
      getCharacterDictionarySelection: async () => snapshot,
      setCharacterDictionarySelection: async (mediaId) => {
        savedIds.push(mediaId);
        return {
          ok: true,
          seriesKey: snapshot.seriesKey,
          selected: snapshot.candidates[0]!,
          staleMediaIds: [10607],
        };
      },
    },
    async ({ modal, state, dom }) => {
      modal.wireDomEvents();

      await modal.openCharacterDictionaryModal();
      assert.equal(state.characterDictionaryModalOpen, true);
      assert.equal(dom.overlay.classList.contains('interactive'), true);
      assert.equal(dom.characterDictionaryModal.classList.contains('hidden'), false);
      assert.equal(dom.characterDictionaryCandidates.children.length, 1);

      pressEnter(modal);
      await flushAsyncWork();

      assert.deepEqual(savedIds, [21355]);
      assert.match(dom.characterDictionaryStatus.textContent, /Override saved/);

      dom.characterDictionaryClose.dispatchEvent('click');
      assert.equal(state.characterDictionaryModalOpen, false);
    },
  );
});

test('character dictionary modal shows refresh errors without rejecting open', async () => {
  await withHarness(
    {
      getCharacterDictionarySelection: async () => {
        throw new Error('candidate lookup failed');
      },
    },
    async ({ modal, state, dom }) => {
      await modal.openCharacterDictionaryModal();

      assert.equal(state.characterDictionaryModalOpen, true);
      assert.equal(dom.characterDictionaryStatus.textContent, 'candidate lookup failed');
      assert.equal(dom.characterDictionaryStatus.classList.contains('error'), true);
    },
  );
});

test('character dictionary modal seeds search input and waits for manual search', async () => {
  const initialSnapshot: CharacterDictionarySelectionSnapshot = {
    seriesKey: 'kage-no-jitsuryokusha-ni-naritakute-2022',
    guessTitle: 'Kage no Jitsuryokusha ni Naritakute!',
    current: null,
    override: null,
    candidates: [],
  };
  const searchedSnapshot: CharacterDictionarySelectionSnapshot = {
    ...initialSnapshot,
    candidates: [{ id: 130298, title: 'The Eminence in Shadow', episodes: 20 }],
  };
  const searches: Array<string | undefined> = [];

  await withHarness(
    {
      getCharacterDictionarySelection: async (searchText) => {
        searches.push(searchText);
        return searchText ? searchedSnapshot : initialSnapshot;
      },
    },
    async ({ modal, dom }) => {
      modal.wireDomEvents();

      await modal.openCharacterDictionaryModal();

      assert.deepEqual(searches, ['']);
      assert.equal(
        dom.characterDictionarySearchInput.value,
        'Kage no Jitsuryokusha ni Naritakute!',
      );
      assert.equal(dom.characterDictionaryCandidates.children.length, 1);
      assert.match(dom.characterDictionaryStatus.textContent, /Enter a title/);

      dom.characterDictionarySearchInput.value = 'Eminence in Shadow';
      dom.characterDictionarySearchButton.dispatchEvent('click');
      await flushAsyncWork();

      assert.deepEqual(searches, ['', 'Eminence in Shadow']);
      assert.equal(dom.characterDictionaryCandidates.children.length, 1);
      assert.match(dom.characterDictionaryStatus.textContent, /Select the correct AniList entry/);
    },
  );
});

test('character dictionary modal marks override candidate as selected', async () => {
  const konosuba = {
    id: 21202,
    title: "KONOSUBA -God's blessing on this wonderful world!",
    episodes: 10,
  };

  await withHarness(
    {
      getCharacterDictionarySelection: async () => ({
        seriesKey: 'konosuba-gods-blessing-on-this-wonderful-world-2016',
        guessTitle: "KonoSuba - God's blessing on this wonderful world!",
        current: null,
        override: konosuba,
        candidates: [konosuba],
      }),
    },
    async ({ modal, dom }) => {
      await modal.openCharacterDictionaryModal();

      const item = dom.characterDictionaryCandidates.children[0] as NodeStub;
      const button = item.children[1] as NodeStub;
      assert.equal(button.textContent, 'Selected');
      assert.equal(button.disabled, true);
    },
  );
});

test('character dictionary modal does not resave the active override from keyboard apply', async () => {
  const reZero = { id: 21355, title: 'Re:ZERO -Starting Life in Another World-', episodes: 25 };
  const savedIds: number[] = [];

  await withHarness(
    {
      getCharacterDictionarySelection: async () => ({
        seriesKey: 're-zero-starting-life-in-another-world-2016',
        guessTitle: 'Re ZERO, Starting Life in Another World',
        current: reZero,
        override: reZero,
        candidates: [reZero],
      }),
      setCharacterDictionarySelection: async (mediaId) => {
        savedIds.push(mediaId);
        return { ok: true, seriesKey: '', selected: reZero, staleMediaIds: [] };
      },
    },
    async ({ modal }) => {
      await modal.openCharacterDictionaryModal();
      pressEnter(modal);
      await flushAsyncWork();

      assert.deepEqual(savedIds, []);
    },
  );
});

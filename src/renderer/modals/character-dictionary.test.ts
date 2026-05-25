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

function createElementStub() {
  return {
    className: '',
    textContent: '',
    type: '',
    children: [] as unknown[],
    classList: createClassList(),
    append(...children: unknown[]) {
      this.children.push(...children);
    },
    addEventListener: () => {},
  };
}

function createNodeStub(hidden = false) {
  const listeners = new Map<string, Array<() => void>>();
  return {
    textContent: '',
    children: [] as unknown[],
    classList: createClassList(hidden ? ['hidden'] : []),
    setAttribute: () => {},
    addEventListener: (event: string, listener: () => void) => {
      listeners.set(event, [...(listeners.get(event) ?? []), listener]);
    },
    dispatchEvent: (event: string) => {
      for (const listener of listeners.get(event) ?? []) listener();
    },
    replaceChildren(...children: unknown[]) {
      this.children = [...children];
    },
  };
}

function flushAsyncWork(): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, 0);
  });
}

test('character dictionary modal announces open before AniList refresh resolves', async () => {
  const previousWindow = globalThis.window;
  const previousDocument = globalThis.document;
  let resolveSelection: (snapshot: CharacterDictionarySelectionSnapshot) => void = () => {};
  const selectionPromise = new Promise<CharacterDictionarySelectionSnapshot>((resolve) => {
    resolveSelection = resolve;
  });
  const events: string[] = [];
  const overlay = createNodeStub();
  const modalNode = createNodeStub(true);
  const state = createRendererState();

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        getCharacterDictionarySelection: () => selectionPromise,
        setCharacterDictionarySelection: async () => ({
          ok: false,
          seriesKey: 'test',
          selected: { id: 0, title: '', episodes: null },
          staleMediaIds: [],
        }),
        notifyOverlayModalClosed: () => {},
        notifyOverlayModalOpened: (modal: string) => {
          events.push(`notify:${modal}`);
        },
      } satisfies Pick<
        ElectronAPI,
        | 'getCharacterDictionarySelection'
        | 'setCharacterDictionarySelection'
        | 'notifyOverlayModalClosed'
        | 'notifyOverlayModalOpened'
      >,
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createElementStub(),
    },
  });

  try {
    const modal = createCharacterDictionaryModal(
      {
        state,
        dom: {
          overlay,
          characterDictionaryModal: modalNode,
          characterDictionaryClose: createNodeStub(),
          characterDictionarySummary: createNodeStub(),
          characterDictionaryCurrent: createNodeStub(),
          characterDictionaryCandidates: createNodeStub(),
          characterDictionaryStatus: createNodeStub(),
        },
      } as never,
      {
        modalStateReader: { isAnyModalOpen: () => false },
        syncSettingsModalSubtitleSuppression: () => {
          events.push('sync-subtitle-suppression');
        },
      },
    );

    const openPromise = modal.openCharacterDictionaryModal();

    assert.equal(state.characterDictionaryModalOpen, true);
    assert.equal(modalNode.classList.contains('hidden'), false);
    assert.deepEqual(events, ['sync-subtitle-suppression', 'notify:character-dictionary']);

    resolveSelection({
      seriesKey: 'tower-of-god-2020',
      guessTitle: 'Tower of God',
      current: null,
      override: null,
      candidates: [{ id: 115230, title: 'Tower of God', episodes: 13 }],
    });
    await openPromise;
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('character dictionary modal loads candidates and applies selected override', async () => {
  const previousWindow = globalThis.window;
  const previousDocument = globalThis.document;
  const snapshot: CharacterDictionarySelectionSnapshot = {
    seriesKey: 're-zero-starting-life-in-another-world-2016',
    guessTitle: 'Re ZERO, Starting Life in Another World',
    current: { id: 10607, title: 'Rerere no Tensai Bakabon', episodes: 24 },
    override: null,
    candidates: [{ id: 21355, title: 'Re:ZERO -Starting Life in Another World-', episodes: 25 }],
  };
  const calls: number[] = [];
  const overlay = createNodeStub();
  const modalNode = createNodeStub(true);
  const closeButton = createNodeStub();
  const candidates = createNodeStub();
  const status = createNodeStub();
  const state = createRendererState();

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        getCharacterDictionarySelection: async () => snapshot,
        setCharacterDictionarySelection: async (mediaId: number) => {
          calls.push(mediaId);
          return {
            ok: true,
            seriesKey: snapshot.seriesKey,
            selected: snapshot.candidates[0]!,
            staleMediaIds: [10607],
          };
        },
        notifyOverlayModalClosed: () => {},
        notifyOverlayModalOpened: () => {},
      } satisfies Pick<
        ElectronAPI,
        | 'getCharacterDictionarySelection'
        | 'setCharacterDictionarySelection'
        | 'notifyOverlayModalClosed'
        | 'notifyOverlayModalOpened'
      >,
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createElementStub(),
    },
  });

  try {
    const modal = createCharacterDictionaryModal(
      {
        state,
        dom: {
          overlay,
          characterDictionaryModal: modalNode,
          characterDictionaryClose: closeButton,
          characterDictionarySummary: createNodeStub(),
          characterDictionaryCurrent: createNodeStub(),
          characterDictionaryCandidates: candidates,
          characterDictionaryStatus: status,
        },
      } as never,
      {
        modalStateReader: { isAnyModalOpen: () => false },
        syncSettingsModalSubtitleSuppression: () => {},
      },
    );
    modal.wireDomEvents();

    await modal.openCharacterDictionaryModal();
    assert.equal(state.characterDictionaryModalOpen, true);
    assert.equal(overlay.classList.contains('interactive'), true);
    assert.equal(modalNode.classList.contains('hidden'), false);
    assert.equal(candidates.children.length, 1);

    modal.handleCharacterDictionaryKeydown({
      key: 'Enter',
      preventDefault: () => {},
    } as KeyboardEvent);
    await flushAsyncWork();

    assert.deepEqual(calls, [21355]);
    assert.match(status.textContent, /Override saved/);

    closeButton.dispatchEvent('click');
    assert.equal(state.characterDictionaryModalOpen, false);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('character dictionary modal shows refresh errors without rejecting open', async () => {
  const previousWindow = globalThis.window;
  const overlay = createNodeStub();
  const modalNode = createNodeStub(true);
  const status = createNodeStub();
  const state = createRendererState();

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        getCharacterDictionarySelection: async () => {
          throw new Error('candidate lookup failed');
        },
        setCharacterDictionarySelection: async () => ({
          ok: false,
          seriesKey: 'test',
          selected: { id: 0, title: '', episodes: null },
          staleMediaIds: [],
        }),
        notifyOverlayModalClosed: () => {},
        notifyOverlayModalOpened: () => {},
      } satisfies Pick<
        ElectronAPI,
        | 'getCharacterDictionarySelection'
        | 'setCharacterDictionarySelection'
        | 'notifyOverlayModalClosed'
        | 'notifyOverlayModalOpened'
      >,
    },
  });

  try {
    const modal = createCharacterDictionaryModal(
      {
        state,
        dom: {
          overlay,
          characterDictionaryModal: modalNode,
          characterDictionaryClose: createNodeStub(),
          characterDictionarySummary: createNodeStub(),
          characterDictionaryCurrent: createNodeStub(),
          characterDictionaryCandidates: createNodeStub(),
          characterDictionaryStatus: status,
        },
      } as never,
      {
        modalStateReader: { isAnyModalOpen: () => false },
        syncSettingsModalSubtitleSuppression: () => {},
      },
    );

    await modal.openCharacterDictionaryModal();

    assert.equal(state.characterDictionaryModalOpen, true);
    assert.equal(status.textContent, 'candidate lookup failed');
    assert.equal(status.classList.contains('error'), true);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
  }
});

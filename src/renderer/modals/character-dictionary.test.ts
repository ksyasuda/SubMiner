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
  return {
    textContent: '',
    children: [] as unknown[],
    classList: createClassList(hidden ? ['hidden'] : []),
    setAttribute: () => {},
    addEventListener: () => {},
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
      } satisfies Pick<
        ElectronAPI,
        | 'getCharacterDictionarySelection'
        | 'setCharacterDictionarySelection'
        | 'notifyOverlayModalClosed'
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
          characterDictionaryCandidates: candidates,
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
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

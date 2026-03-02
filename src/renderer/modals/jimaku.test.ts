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

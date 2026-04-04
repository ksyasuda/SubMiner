import assert from 'node:assert/strict';
import test from 'node:test';

import type { ElectronAPI, PlaylistBrowserSnapshot } from '../../types';
import { createRendererState } from '../state.js';
import { createPlaylistBrowserModal } from './playlist-browser.js';

function createClassList(initialTokens: string[] = []) {
  const tokens = new Set(initialTokens);
  return {
    add: (...entries: string[]) => {
      for (const entry of entries) tokens.add(entry);
    },
    remove: (...entries: string[]) => {
      for (const entry of entries) tokens.delete(entry);
    },
    contains: (entry: string) => tokens.has(entry),
    toggle: (entry: string, force?: boolean) => {
      if (force === true) tokens.add(entry);
      else if (force === false) tokens.delete(entry);
      else if (tokens.has(entry)) tokens.delete(entry);
      else tokens.add(entry);
    },
  };
}

function createFakeElement() {
  const attributes = new Map<string, string>();
  return {
    textContent: '',
    innerHTML: '',
    children: [] as unknown[],
    listeners: new Map<string, Array<(event?: unknown) => void>>(),
    classList: createClassList(['hidden']),
    appendChild(child: unknown) {
      this.children.push(child);
      return child;
    },
    append(...children: unknown[]) {
      this.children.push(...children);
    },
    replaceChildren(...children: unknown[]) {
      this.children = [...children];
    },
    addEventListener(type: string, listener: (event?: unknown) => void) {
      const bucket = this.listeners.get(type) ?? [];
      bucket.push(listener);
      this.listeners.set(type, bucket);
    },
    setAttribute(name: string, value: string) {
      attributes.set(name, value);
    },
    getAttribute(name: string) {
      return attributes.get(name) ?? null;
    },
    focus() {},
  };
}

function createPlaylistRow() {
  return {
    className: '',
    classList: createClassList(),
    dataset: {} as Record<string, string>,
    textContent: '',
    children: [] as unknown[],
    listeners: new Map<string, Array<(event?: unknown) => void>>(),
    append(...children: unknown[]) {
      this.children.push(...children);
    },
    appendChild(child: unknown) {
      this.children.push(child);
      return child;
    },
    addEventListener(type: string, listener: (event?: unknown) => void) {
      const bucket = this.listeners.get(type) ?? [];
      bucket.push(listener);
      this.listeners.set(type, bucket);
    },
    setAttribute() {},
  };
}

function createListStub() {
  return {
    innerHTML: '',
    children: [] as ReturnType<typeof createPlaylistRow>[],
    appendChild(child: ReturnType<typeof createPlaylistRow>) {
      this.children.push(child);
      return child;
    },
    replaceChildren(...children: ReturnType<typeof createPlaylistRow>[]) {
      this.children = [...children];
    },
  };
}

function createSnapshot(): PlaylistBrowserSnapshot {
  return {
    directoryPath: '/tmp/show',
    directoryAvailable: true,
    directoryStatus: '/tmp/show',
    currentFilePath: '/tmp/show/Show - S01E02.mkv',
    playingIndex: 1,
    directoryItems: [
      {
        path: '/tmp/show/Show - S01E01.mkv',
        basename: 'Show - S01E01.mkv',
        episodeLabel: 'S1E1',
        isCurrentFile: false,
      },
      {
        path: '/tmp/show/Show - S01E02.mkv',
        basename: 'Show - S01E02.mkv',
        episodeLabel: 'S1E2',
        isCurrentFile: true,
      },
    ],
    playlistItems: [
      {
        index: 0,
        id: 1,
        filename: '/tmp/show/Show - S01E01.mkv',
        title: 'Episode 1',
        displayLabel: 'Episode 1',
        current: false,
        playing: false,
        path: '/tmp/show/Show - S01E01.mkv',
      },
      {
        index: 1,
        id: 2,
        filename: '/tmp/show/Show - S01E02.mkv',
        title: 'Episode 2',
        displayLabel: 'Episode 2',
        current: true,
        playing: true,
        path: '/tmp/show/Show - S01E02.mkv',
      },
    ],
  };
}

function createMutationSnapshot(): PlaylistBrowserSnapshot {
  return {
    directoryPath: '/tmp/show',
    directoryAvailable: true,
    directoryStatus: '/tmp/show',
    currentFilePath: '/tmp/show/Show - S01E02.mkv',
    playingIndex: 0,
    directoryItems: [
      {
        path: '/tmp/show/Show - S01E01.mkv',
        basename: 'Show - S01E01.mkv',
        episodeLabel: 'S1E1',
        isCurrentFile: false,
      },
      {
        path: '/tmp/show/Show - S01E02.mkv',
        basename: 'Show - S01E02.mkv',
        episodeLabel: 'S1E2',
        isCurrentFile: true,
      },
      {
        path: '/tmp/show/Show - S01E03.mkv',
        basename: 'Show - S01E03.mkv',
        episodeLabel: 'S1E3',
        isCurrentFile: false,
      },
    ],
    playlistItems: [
      {
        index: 1,
        id: 2,
        filename: '/tmp/show/Show - S01E02.mkv',
        title: 'Episode 2',
        displayLabel: 'Episode 2',
        current: true,
        playing: true,
        path: '/tmp/show/Show - S01E02.mkv',
      },
      {
        index: 2,
        id: 3,
        filename: '/tmp/show/Show - S01E03.mkv',
        title: 'Episode 3',
        displayLabel: 'Episode 3',
        current: false,
        playing: false,
        path: '/tmp/show/Show - S01E03.mkv',
      },
      {
        index: 0,
        id: 1,
        filename: '/tmp/show/Show - S01E01.mkv',
        title: 'Episode 1',
        displayLabel: 'Episode 1',
        current: false,
        playing: false,
        path: '/tmp/show/Show - S01E01.mkv',
      },
    ],
  };
}

function restoreGlobalDescriptor<K extends keyof typeof globalThis>(
  key: K,
  descriptor: PropertyDescriptor | undefined,
) {
  if (descriptor) {
    Object.defineProperty(globalThis, key, descriptor);
    return;
  }
  Reflect.deleteProperty(globalThis, key);
}

function createPlaylistBrowserDomFixture() {
  return {
    overlay: {
      classList: createClassList(),
      focus: () => {},
    },
    playlistBrowserModal: createFakeElement(),
    playlistBrowserTitle: createFakeElement(),
    playlistBrowserStatus: createFakeElement(),
    playlistBrowserDirectoryList: createListStub(),
    playlistBrowserPlaylistList: createListStub(),
    playlistBrowserClose: createFakeElement(),
  };
}

function createPlaylistBrowserElectronApi(overrides?: Partial<ElectronAPI>): ElectronAPI {
  return {
    getPlaylistBrowserSnapshot: async () => createSnapshot(),
    notifyOverlayModalOpened: () => {},
    notifyOverlayModalClosed: () => {},
    focusMainWindow: async () => {},
    setIgnoreMouseEvents: () => {},
    appendPlaylistBrowserFile: async () => ({
      ok: true,
      message: 'ok',
      snapshot: createSnapshot(),
    }),
    playPlaylistBrowserIndex: async () => ({ ok: true, message: 'ok', snapshot: createSnapshot() }),
    removePlaylistBrowserIndex: async () => ({
      ok: true,
      message: 'ok',
      snapshot: createSnapshot(),
    }),
    movePlaylistBrowserIndex: async () => ({ ok: true, message: 'ok', snapshot: createSnapshot() }),
    ...overrides,
  } as ElectronAPI;
}

function setupPlaylistBrowserModalTest(options?: {
  electronApi?: Partial<ElectronAPI>;
  shouldToggleMouseIgnore?: boolean;
}) {
  const previousWindowDescriptor = Object.getOwnPropertyDescriptor(globalThis, 'window');
  const previousDocumentDescriptor = Object.getOwnPropertyDescriptor(globalThis, 'document');
  const state = createRendererState();
  const dom = createPlaylistBrowserDomFixture();
  const ctx = {
    state,
    platform: {
      shouldToggleMouseIgnore: options?.shouldToggleMouseIgnore ?? false,
    },
    dom,
  };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: createPlaylistBrowserElectronApi(options?.electronApi),
      focus: () => {},
    } satisfies { electronAPI: ElectronAPI; focus: () => void },
    writable: true,
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createPlaylistRow(),
    },
    writable: true,
  });

  return {
    state,
    dom,
    createModal(overrides: Partial<Parameters<typeof createPlaylistBrowserModal>[1]> = {}) {
      return createPlaylistBrowserModal(ctx as never, {
        modalStateReader: { isAnyModalOpen: () => false },
        syncSettingsModalSubtitleSuppression: () => {},
        ...overrides,
      });
    },
    restore() {
      restoreGlobalDescriptor('window', previousWindowDescriptor);
      restoreGlobalDescriptor('document', previousDocumentDescriptor);
    },
  };
}

test('playlist browser test cleanup must delete injected globals that were originally absent', () => {
  const previousWindowDescriptor = Object.getOwnPropertyDescriptor(globalThis, 'window');
  const previousDocumentDescriptor = Object.getOwnPropertyDescriptor(globalThis, 'document');
  try {
    Reflect.deleteProperty(globalThis, 'window');
    Reflect.deleteProperty(globalThis, 'document');
    assert.equal(Object.prototype.hasOwnProperty.call(globalThis, 'window'), false);
    assert.equal(Object.prototype.hasOwnProperty.call(globalThis, 'document'), false);

    const env = setupPlaylistBrowserModalTest();

    assert.equal(Object.prototype.hasOwnProperty.call(globalThis, 'window'), true);
    assert.equal(Object.prototype.hasOwnProperty.call(globalThis, 'document'), true);
    env.restore();

    assert.equal(Object.prototype.hasOwnProperty.call(globalThis, 'window'), false);
    assert.equal(Object.prototype.hasOwnProperty.call(globalThis, 'document'), false);
    assert.equal(typeof globalThis.window, 'undefined');
    assert.equal(typeof globalThis.document, 'undefined');
  } finally {
    restoreGlobalDescriptor('window', previousWindowDescriptor);
    restoreGlobalDescriptor('document', previousDocumentDescriptor);
  }
});

test('playlist browser modal opens with playlist-focused current item selection', async () => {
  const notifications: string[] = [];
  const env = setupPlaylistBrowserModalTest({
    electronApi: {
      notifyOverlayModalOpened: (modal: string) => notifications.push(`open:${modal}`),
      notifyOverlayModalClosed: (modal: string) => notifications.push(`close:${modal}`),
    },
  });

  try {
    const modal = env.createModal();

    await modal.openPlaylistBrowserModal();

    assert.equal(env.state.playlistBrowserModalOpen, true);
    assert.equal(env.state.playlistBrowserActivePane, 'playlist');
    assert.equal(env.state.playlistBrowserSelectedPlaylistIndex, 1);
    assert.equal(env.state.playlistBrowserSelectedDirectoryIndex, 1);
    assert.equal(env.dom.playlistBrowserDirectoryList.children.length, 2);
    assert.equal(env.dom.playlistBrowserPlaylistList.children.length, 2);
    assert.equal(env.dom.playlistBrowserDirectoryList.children[0]?.children.length, 2);
    assert.equal(env.dom.playlistBrowserPlaylistList.children[0]?.children.length, 2);
    assert.deepEqual(notifications, ['open:playlist-browser']);
  } finally {
    env.restore();
  }
});

test('playlist browser modal action buttons stop double-click propagation', async () => {
  const env = setupPlaylistBrowserModalTest();

  try {
    const modal = env.createModal();

    await modal.openPlaylistBrowserModal();

    const row = env.dom.playlistBrowserDirectoryList.children[0] as
      | ReturnType<typeof createPlaylistRow>
      | undefined;
    const trailing = row?.children?.[1] as ReturnType<typeof createPlaylistRow> | undefined;
    const button = trailing?.children?.at(-1) as
      | { listeners?: Map<string, Array<(event?: unknown) => void>> }
      | undefined;
    const dblclickHandler = button?.listeners?.get('dblclick')?.[0];

    assert.equal(typeof dblclickHandler, 'function');
    let stopped = false;
    dblclickHandler?.({
      stopPropagation: () => {
        stopped = true;
      },
    });

    assert.equal(stopped, true);
  } finally {
    env.restore();
  }
});

test('playlist browser preserves prior selection across mutation snapshots', async () => {
  const env = setupPlaylistBrowserModalTest({
    electronApi: {
      getPlaylistBrowserSnapshot: async () => ({
        ...createSnapshot(),
        directoryItems: [
          ...createSnapshot().directoryItems,
          {
            path: '/tmp/show/Show - S01E03.mkv',
            basename: 'Show - S01E03.mkv',
            episodeLabel: 'S1E3',
            isCurrentFile: false,
          },
        ],
        playlistItems: [
          ...createSnapshot().playlistItems,
          {
            index: 2,
            id: 3,
            filename: '/tmp/show/Show - S01E03.mkv',
            title: 'Episode 3',
            displayLabel: 'Episode 3',
            current: false,
            playing: false,
            path: '/tmp/show/Show - S01E03.mkv',
          },
        ],
      }),
      appendPlaylistBrowserFile: async () => ({
        ok: true,
        message: 'Queued file',
        snapshot: createMutationSnapshot(),
      }),
    },
  });

  try {
    const modal = env.createModal();

    await modal.openPlaylistBrowserModal();
    env.state.playlistBrowserActivePane = 'directory';
    env.state.playlistBrowserSelectedDirectoryIndex = 2;
    env.state.playlistBrowserSelectedPlaylistIndex = 0;

    await modal.handlePlaylistBrowserKeydown({
      key: 'Enter',
      code: 'Enter',
      preventDefault: () => {},
      ctrlKey: false,
      metaKey: false,
      shiftKey: false,
    } as never);

    assert.equal(env.state.playlistBrowserSelectedDirectoryIndex, 2);
    assert.equal(env.state.playlistBrowserSelectedPlaylistIndex, 2);
  } finally {
    env.restore();
  }
});

test('playlist browser modal keydown routes append, remove, reorder, tab switch, and play', async () => {
  const calls: Array<[string, unknown[]]> = [];
  const notifications: string[] = [];
  const env = setupPlaylistBrowserModalTest({
    electronApi: {
      notifyOverlayModalOpened: (modal: string) => notifications.push(`open:${modal}`),
      notifyOverlayModalClosed: (modal: string) => notifications.push(`close:${modal}`),
      appendPlaylistBrowserFile: async (filePath: string) => {
        calls.push(['append', [filePath]]);
        return { ok: true, message: 'append-ok', snapshot: createMutationSnapshot() };
      },
      playPlaylistBrowserIndex: async (index: number) => {
        calls.push(['play', [index]]);
        return { ok: true, message: 'play-ok', snapshot: createSnapshot() };
      },
      removePlaylistBrowserIndex: async (index: number) => {
        calls.push(['remove', [index]]);
        return { ok: true, message: 'remove-ok', snapshot: createSnapshot() };
      },
      movePlaylistBrowserIndex: async (index: number, direction: -1 | 1) => {
        calls.push(['move', [index, direction]]);
        return { ok: true, message: 'move-ok', snapshot: createSnapshot() };
      },
    },
  });

  try {
    const modal = env.createModal();

    await modal.openPlaylistBrowserModal();

    const preventDefault = () => {};
    env.state.playlistBrowserActivePane = 'directory';
    env.state.playlistBrowserSelectedDirectoryIndex = 0;
    await modal.handlePlaylistBrowserKeydown({
      key: 'Enter',
      code: 'Enter',
      preventDefault,
      ctrlKey: false,
      metaKey: false,
      shiftKey: false,
    } as never);

    await modal.handlePlaylistBrowserKeydown({
      key: 'Tab',
      code: 'Tab',
      preventDefault,
      ctrlKey: false,
      metaKey: false,
      shiftKey: false,
    } as never);
    assert.equal(env.state.playlistBrowserActivePane, 'playlist');

    await modal.handlePlaylistBrowserKeydown({
      key: 'ArrowDown',
      code: 'ArrowDown',
      preventDefault,
      ctrlKey: true,
      metaKey: false,
      shiftKey: false,
    } as never);

    await modal.handlePlaylistBrowserKeydown({
      key: 'Delete',
      code: 'Delete',
      preventDefault,
      ctrlKey: false,
      metaKey: false,
      shiftKey: false,
    } as never);

    await modal.handlePlaylistBrowserKeydown({
      key: 'Enter',
      code: 'Enter',
      preventDefault,
      ctrlKey: false,
      metaKey: false,
      shiftKey: false,
    } as never);

    assert.deepEqual(calls, [
      ['append', ['/tmp/show/Show - S01E01.mkv']],
      ['move', [1, 1]],
      ['remove', [1]],
      ['play', [1]],
    ]);
    assert.equal(env.state.playlistBrowserModalOpen, false);
    assert.deepEqual(notifications, ['open:playlist-browser', 'close:playlist-browser']);
  } finally {
    env.restore();
  }
});

test('playlist browser keeps modal open when playing selected queue item fails', async () => {
  const notifications: string[] = [];
  const env = setupPlaylistBrowserModalTest({
    electronApi: {
      notifyOverlayModalOpened: (modal: string) => notifications.push(`open:${modal}`),
      notifyOverlayModalClosed: (modal: string) => notifications.push(`close:${modal}`),
      playPlaylistBrowserIndex: async () => ({ ok: false, message: 'play failed' }),
    },
  });

  try {
    const modal = env.createModal();

    await modal.openPlaylistBrowserModal();
    assert.equal(env.state.playlistBrowserModalOpen, true);

    await modal.handlePlaylistBrowserKeydown({
      key: 'Enter',
      code: 'Enter',
      preventDefault: () => {},
      ctrlKey: false,
      metaKey: false,
      shiftKey: false,
    } as never);

    assert.equal(env.state.playlistBrowserModalOpen, true);
    assert.equal(env.dom.playlistBrowserStatus.textContent, 'play failed');
    assert.equal(env.dom.playlistBrowserStatus.classList.contains('error'), true);
    assert.deepEqual(notifications, ['open:playlist-browser']);
  } finally {
    env.restore();
  }
});

test('playlist browser refresh failure clears stale rendered rows and reports the error', async () => {
  const notifications: string[] = [];
  let refreshShouldFail = false;
  const env = setupPlaylistBrowserModalTest({
    electronApi: {
      getPlaylistBrowserSnapshot: async () => {
        if (refreshShouldFail) {
          throw new Error('snapshot failed');
        }
        return createSnapshot();
      },
      notifyOverlayModalOpened: (modal: string) => notifications.push(`open:${modal}`),
      notifyOverlayModalClosed: (modal: string) => notifications.push(`close:${modal}`),
    },
  });

  try {
    const modal = env.createModal();

    await modal.openPlaylistBrowserModal();
    assert.equal(env.dom.playlistBrowserDirectoryList.children.length, 2);
    assert.equal(env.dom.playlistBrowserPlaylistList.children.length, 2);

    refreshShouldFail = true;
    await modal.refreshSnapshot();

    assert.equal(env.state.playlistBrowserSnapshot, null);
    assert.equal(env.dom.playlistBrowserDirectoryList.children.length, 0);
    assert.equal(env.dom.playlistBrowserPlaylistList.children.length, 0);
    assert.equal(env.dom.playlistBrowserTitle.textContent, 'Playlist Browser');
    assert.equal(env.dom.playlistBrowserStatus.textContent, 'snapshot failed');
    assert.equal(env.dom.playlistBrowserStatus.classList.contains('error'), true);
    assert.deepEqual(notifications, ['open:playlist-browser']);
  } finally {
    env.restore();
  }
});

test('playlist browser close clears rendered snapshot ui', async () => {
  const notifications: string[] = [];
  const env = setupPlaylistBrowserModalTest({
    electronApi: {
      notifyOverlayModalOpened: (modal: string) => notifications.push(`open:${modal}`),
      notifyOverlayModalClosed: (modal: string) => notifications.push(`close:${modal}`),
    },
  });

  try {
    const modal = env.createModal();

    await modal.openPlaylistBrowserModal();
    assert.equal(env.dom.playlistBrowserDirectoryList.children.length, 2);
    assert.equal(env.dom.playlistBrowserPlaylistList.children.length, 2);

    modal.closePlaylistBrowserModal();

    assert.equal(env.state.playlistBrowserSnapshot, null);
    assert.equal(env.state.playlistBrowserStatus, '');
    assert.equal(env.dom.playlistBrowserDirectoryList.children.length, 0);
    assert.equal(env.dom.playlistBrowserPlaylistList.children.length, 0);
    assert.equal(env.dom.playlistBrowserTitle.textContent, 'Playlist Browser');
    assert.equal(env.dom.playlistBrowserStatus.textContent, '');
    assert.deepEqual(notifications, ['open:playlist-browser', 'close:playlist-browser']);
  } finally {
    env.restore();
  }
});

test('playlist browser open is ignored while another modal is already open', async () => {
  const notifications: string[] = [];
  let snapshotCalls = 0;
  const env = setupPlaylistBrowserModalTest({
    electronApi: {
      getPlaylistBrowserSnapshot: async () => {
        snapshotCalls += 1;
        return createSnapshot();
      },
      notifyOverlayModalOpened: (modal: string) => notifications.push(`open:${modal}`),
      notifyOverlayModalClosed: (modal: string) => notifications.push(`close:${modal}`),
    },
  });

  try {
    const modal = env.createModal({
      modalStateReader: { isAnyModalOpen: () => true },
    });

    await modal.openPlaylistBrowserModal();

    assert.equal(env.state.playlistBrowserModalOpen, false);
    assert.equal(snapshotCalls, 0);
    assert.equal(env.dom.overlay.classList.contains('interactive'), false);
    assert.deepEqual(notifications, []);
  } finally {
    env.restore();
  }
});

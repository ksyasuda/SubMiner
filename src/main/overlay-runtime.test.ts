import assert from 'node:assert/strict';
import test from 'node:test';
import { createOverlayModalRuntimeService } from './overlay-runtime';

type MockWindow = {
  destroyed: boolean;
  visible: boolean;
  focused: boolean;
  ignoreMouseEvents: boolean;
  forwardedIgnoreMouseEvents: boolean;
  webContentsFocused: boolean;
  showCount: number;
  hideCount: number;
  sent: unknown[][];
  loading: boolean;
  url: string;
  contentReady: boolean;
  loadCallbacks: Array<() => void>;
  readyToShowCallbacks: Array<() => void>;
};

function createMockWindow(): MockWindow & {
  isDestroyed: () => boolean;
  isVisible: () => boolean;
  isFocused: () => boolean;
  getURL: () => string;
  setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => void;
  setAlwaysOnTop: (flag: boolean, level?: string, relativeLevel?: number) => void;
  moveTop: () => void;
  getShowCount: () => number;
  getHideCount: () => number;
  show: () => void;
  hide: () => void;
  destroy: () => void;
  focus: () => void;
  emitDidFinishLoad: () => void;
  emitReadyToShow: () => void;
  once: (event: 'ready-to-show', cb: () => void) => void;
  webContents: {
    focused: boolean;
    isLoading: () => boolean;
    getURL: () => string;
    send: (channel: string, payload?: unknown) => void;
    isFocused: () => boolean;
    once: (event: 'did-finish-load', cb: () => void) => void;
    focus: () => void;
  };
} {
  const state: MockWindow = {
    destroyed: false,
    visible: false,
    focused: false,
    ignoreMouseEvents: false,
    forwardedIgnoreMouseEvents: false,
    webContentsFocused: false,
    showCount: 0,
    hideCount: 0,
    sent: [],
    loading: false,
    url: 'file:///overlay/index.html?layer=modal',
    contentReady: true,
    loadCallbacks: [],
    readyToShowCallbacks: [],
  };
  const window = {
    ...state,
    isDestroyed: () => state.destroyed,
    isVisible: () => state.visible,
    isFocused: () => state.focused,
    getURL: () => state.url,
    setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
      state.ignoreMouseEvents = ignore;
      state.forwardedIgnoreMouseEvents = options?.forward === true;
    },
    setAlwaysOnTop: (_flag: boolean, _level?: string, _relativeLevel?: number) => {},
    moveTop: () => {},
    getShowCount: () => state.showCount,
    getHideCount: () => state.hideCount,
    show: () => {
      state.visible = true;
      state.showCount += 1;
    },
    hide: () => {
      state.visible = false;
      state.hideCount += 1;
    },
    destroy: () => {
      state.destroyed = true;
      state.visible = false;
    },
    focus: () => {
      state.focused = true;
    },
    emitDidFinishLoad: () => {
      const callbacks = state.loadCallbacks.splice(0);
      for (const callback of callbacks) {
        callback();
      }
    },
    emitReadyToShow: () => {
      const callbacks = state.readyToShowCallbacks.splice(0);
      for (const callback of callbacks) {
        callback();
      }
    },
    once: (_event: 'ready-to-show', cb: () => void) => {
      state.readyToShowCallbacks.push(cb);
    },
    webContents: {
      isLoading: () => state.loading,
      getURL: () => state.url,
      send: (channel: string, payload?: unknown) => {
        if (payload === undefined) {
          state.sent.push([channel]);
          return;
        }
        state.sent.push([channel, payload]);
      },
      focused: false,
      isFocused: () => state.webContentsFocused,
      once: (_event: 'did-finish-load', cb: () => void) => {
        state.loadCallbacks.push(cb);
      },
      focus: () => {
        state.webContentsFocused = true;
      },
    },
  };

  Object.defineProperty(window, 'loading', {
    get: () => state.loading,
    set: (value: boolean) => {
      state.loading = value;
    },
  });

  Object.defineProperty(window, 'visible', {
    get: () => state.visible,
    set: (value: boolean) => {
      state.visible = value;
    },
  });

  Object.defineProperty(window, 'focused', {
    get: () => state.focused,
    set: (value: boolean) => {
      state.focused = value;
    },
  });

  Object.defineProperty(window, 'webContentsFocused', {
    get: () => state.webContentsFocused,
    set: (value: boolean) => {
      state.webContentsFocused = value;
    },
  });

  Object.defineProperty(window, 'url', {
    get: () => state.url,
    set: (value: string) => {
      state.url = value;
    },
  });

  Object.defineProperty(window, 'ignoreMouseEvents', {
    get: () => state.ignoreMouseEvents,
    set: (value: boolean) => {
      state.ignoreMouseEvents = value;
    },
  });

  Object.defineProperty(window, 'forwardedIgnoreMouseEvents', {
    get: () => state.forwardedIgnoreMouseEvents,
    set: (value: boolean) => {
      state.forwardedIgnoreMouseEvents = value;
    },
  });

  Object.defineProperty(window, 'contentReady', {
    get: () => state.contentReady,
    set: (value: boolean) => {
      state.contentReady = value;
      (window as typeof window & { __subminerOverlayContentReady?: boolean }).__subminerOverlayContentReady =
        value;
    },
  });

  (window as typeof window & { __subminerOverlayContentReady?: boolean }).__subminerOverlayContentReady =
    state.contentReady;

  return window;
}

test('sendToActiveOverlayWindow targets modal window with full geometry and tracks close restore', () => {
  const window = createMockWindow();
  const calls: string[] = [];
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => null,
    getModalWindow: () => window as never,
    createModalWindow: () => {
      calls.push('create-modal-window');
      return window as never;
    },
    getModalGeometry: () => ({ x: 10, y: 20, width: 300, height: 200 }),
    setModalWindowBounds: (geometry) => {
      calls.push(`bounds:${geometry.x},${geometry.y},${geometry.width},${geometry.height}`);
    },
  });

  const sent = runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });
  assert.equal(sent, true);
  assert.equal(runtime.getRestoreVisibleOverlayOnModalClose().has('runtime-options'), true);
  assert.deepEqual(calls, ['bounds:10,20,300,200']);
  assert.equal(window.getShowCount(), 0);
  runtime.notifyOverlayModalOpened('runtime-options');
  assert.equal(window.getShowCount(), 1);
  assert.equal(window.isFocused(), true);
  assert.deepEqual(window.sent, [['runtime-options:open']]);
});

test('sendToActiveOverlayWindow creates modal window lazily when absent', () => {
  const window = createMockWindow();
  let modalWindow: ReturnType<typeof createMockWindow> | null = null;
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => null,
    getModalWindow: () => modalWindow as never,
    createModalWindow: () => {
      modalWindow = window;
      return modalWindow as never;
    },
    getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
    setModalWindowBounds: () => {},
  });

  assert.equal(
    runtime.sendToActiveOverlayWindow('jimaku:open', undefined, { restoreOnModalClose: 'jimaku' }),
    true,
  );
  assert.equal(window.getShowCount(), 0);
  runtime.notifyOverlayModalOpened('jimaku');
  assert.equal(window.getShowCount(), 1);
  assert.deepEqual(window.sent, [['jimaku:open']]);
});

test('sendToActiveOverlayWindow does not retain restore state when modal creation fails', () => {
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => null,
    getModalWindow: () => null,
    createModalWindow: () => null,
    getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
    setModalWindowBounds: () => {},
  });

  assert.equal(
    runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
      restoreOnModalClose: 'runtime-options',
    }),
    false,
  );
  assert.equal(runtime.getRestoreVisibleOverlayOnModalClose().has('runtime-options'), false);
});

test('sendToActiveOverlayWindow waits for blank modal URL before sending open command', () => {
  const window = createMockWindow();
  window.url = '';
  window.loading = true;
  window.contentReady = false;
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => null,
    getModalWindow: () => window as never,
    createModalWindow: () => {
      throw new Error('modal window should not be created when already present');
    },
    getModalGeometry: () => ({ x: 10, y: 20, width: 300, height: 200 }),
    setModalWindowBounds: () => {},
  });

  const sent = runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });

  assert.equal(sent, true);
  assert.deepEqual(window.sent, []);
  window.loading = false;
  window.url = 'file:///overlay/index.html?layer=modal';
  window.emitDidFinishLoad();
  assert.deepEqual(window.sent, []);

  window.contentReady = true;
  window.emitReadyToShow();

  runtime.notifyOverlayModalOpened('runtime-options');
  assert.deepEqual(window.sent, [['runtime-options:open']]);
  assert.equal(window.getShowCount(), 1);
});

test('handleOverlayModalClosed hides modal window only after all pending modals close', () => {
  const window = createMockWindow();
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => null,
    getModalWindow: () => window as never,
    createModalWindow: () => window as never,
    getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
    setModalWindowBounds: () => {},
  });

  runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });
  runtime.sendToActiveOverlayWindow(
    'subsync:open-manual',
    { sourceTracks: [] },
    {
      restoreOnModalClose: 'subsync',
    },
  );

  runtime.handleOverlayModalClosed('runtime-options');
  assert.equal(window.isDestroyed(), false);

  runtime.handleOverlayModalClosed('subsync');
  assert.equal(window.isDestroyed(), true);
});

test('sendToActiveOverlayWindow prefers visible main overlay window for modal open', () => {
  const mainWindow = createMockWindow();
  mainWindow.visible = true;
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => mainWindow as never,
    getModalWindow: () => null,
    createModalWindow: () => {
      throw new Error('modal window should not be created when main overlay is visible');
    },
    getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
    setModalWindowBounds: () => {},
  });

  const sent = runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });

  assert.equal(sent, true);
  assert.deepEqual(mainWindow.sent, [['runtime-options:open']]);
});

test('sendToActiveOverlayWindow can prefer modal window even when main overlay is visible', () => {
  const mainWindow = createMockWindow();
  mainWindow.visible = true;
  const modalWindow = createMockWindow();
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => mainWindow as never,
    getModalWindow: () => modalWindow as never,
    createModalWindow: () => modalWindow as never,
    getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
    setModalWindowBounds: () => {},
  });

  const sent = runtime.sendToActiveOverlayWindow(
    'youtube:picker-open',
    { sessionId: 'yt-1' },
    {
      restoreOnModalClose: 'youtube-track-picker',
      preferModalWindow: true,
    },
  );

  assert.equal(sent, true);
  assert.deepEqual(mainWindow.sent, []);
  assert.deepEqual(modalWindow.sent, [['youtube:picker-open', { sessionId: 'yt-1' }]]);
});

test('modal window path makes visible main overlay click-through until modal closes', () => {
  const mainWindow = createMockWindow();
  mainWindow.visible = true;
  const modalWindow = createMockWindow();
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => mainWindow as never,
    getModalWindow: () => modalWindow as never,
    createModalWindow: () => modalWindow as never,
    getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
    setModalWindowBounds: () => {},
  });

  const sent = runtime.sendToActiveOverlayWindow(
    'youtube:picker-open',
    { sessionId: 'yt-1' },
    {
      restoreOnModalClose: 'youtube-track-picker',
      preferModalWindow: true,
    },
  );
  runtime.notifyOverlayModalOpened('youtube-track-picker');

  assert.equal(sent, true);
  assert.equal(mainWindow.ignoreMouseEvents, true);
  assert.equal(mainWindow.forwardedIgnoreMouseEvents, true);
  assert.equal(modalWindow.ignoreMouseEvents, false);

  runtime.handleOverlayModalClosed('youtube-track-picker');

  assert.equal(mainWindow.ignoreMouseEvents, true);
});

test('modal window path hides visible main overlay until modal closes', () => {
  const mainWindow = createMockWindow();
  mainWindow.visible = true;
  const modalWindow = createMockWindow();
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => mainWindow as never,
    getModalWindow: () => modalWindow as never,
    createModalWindow: () => modalWindow as never,
    getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
    setModalWindowBounds: () => {},
  });

  runtime.sendToActiveOverlayWindow(
    'youtube:picker-open',
    { sessionId: 'yt-1' },
    {
      restoreOnModalClose: 'youtube-track-picker',
      preferModalWindow: true,
    },
  );
  runtime.notifyOverlayModalOpened('youtube-track-picker');

  assert.equal(mainWindow.getHideCount(), 1);
  assert.equal(mainWindow.isVisible(), false);

  runtime.handleOverlayModalClosed('youtube-track-picker');

  assert.equal(mainWindow.getShowCount(), 0);
  assert.equal(mainWindow.isVisible(), false);
});

test('modal runtime notifies callers when modal input state becomes active/inactive', () => {
  const window = createMockWindow();
  const state: boolean[] = [];
  const runtime = createOverlayModalRuntimeService(
    {
      getMainWindow: () => null,
      getModalWindow: () => window as never,
      createModalWindow: () => window as never,
      getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
      setModalWindowBounds: () => {},
    },
    {
      onModalStateChange: (active: boolean): void => {
        state.push(active);
      },
    },
  );

  runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });
  runtime.sendToActiveOverlayWindow(
    'subsync:open-manual',
    { sourceTracks: [] },
    {
      restoreOnModalClose: 'subsync',
    },
  );
  assert.deepEqual(state, []);
  runtime.notifyOverlayModalOpened('runtime-options');
  assert.deepEqual(state, [true]);

  runtime.handleOverlayModalClosed('runtime-options');
  assert.deepEqual(state, [true]);

  runtime.handleOverlayModalClosed('subsync');
  assert.deepEqual(state, [true, false]);
});

test('notifyOverlayModalOpened enables input on visible main overlay window when no modal window exists', () => {
  const mainWindow = createMockWindow();
  mainWindow.visible = true;
  mainWindow.ignoreMouseEvents = true;
  const state: boolean[] = [];

  const runtime = createOverlayModalRuntimeService(
    {
      getMainWindow: () => mainWindow as never,
      getModalWindow: () => null,
      createModalWindow: () => {
        throw new Error('modal window should not be created when main overlay is visible');
      },
      getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
      setModalWindowBounds: () => {},
    },
    {
      onModalStateChange: (active: boolean): void => {
        state.push(active);
      },
    },
  );

  const sent = runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });
  runtime.notifyOverlayModalOpened('runtime-options');

  assert.equal(sent, true);
  assert.deepEqual(state, [true]);
  assert.equal(mainWindow.ignoreMouseEvents, false);
  assert.equal(mainWindow.isFocused(), true);
  assert.equal(mainWindow.webContentsFocused, true);
});

test('handleOverlayModalClosed is a no-op when no modal window can be targeted', () => {
  const state: boolean[] = [];
  const runtime = createOverlayModalRuntimeService(
    {
      getMainWindow: () => null,
      getModalWindow: () => null,
      createModalWindow: () => null,
      getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
      setModalWindowBounds: () => {},
    },
    {
      onModalStateChange: (active: boolean): void => {
        state.push(active);
      },
    },
  );

  const sent = runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });
  assert.equal(sent, false);
  runtime.notifyOverlayModalOpened('runtime-options');
  runtime.handleOverlayModalClosed('runtime-options');

  assert.deepEqual(state, []);
});

test('handleOverlayModalClosed destroys modal window for single kiku modal', () => {
  const window = createMockWindow();
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => null,
    getModalWindow: () => window as never,
    createModalWindow: () => window as never,
    getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
    setModalWindowBounds: () => {},
  });

  runtime.sendToActiveOverlayWindow(
    'kiku:field-grouping-open',
    { test: true },
    {
      restoreOnModalClose: 'kiku',
    },
  );
  runtime.handleOverlayModalClosed('kiku');

  assert.equal(window.isDestroyed(), true);
  assert.equal(runtime.getRestoreVisibleOverlayOnModalClose().size, 0);
});

test('modal fallback reveal skips showing window when content is not ready', async () => {
  const window = createMockWindow();
  let scheduledReveal: (() => void) | null = null;
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => null,
    getModalWindow: () => window as never,
    createModalWindow: () => {
      throw new Error('modal window should not be created when already present');
    },
    getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
    setModalWindowBounds: () => {},
  }, {
    scheduleRevealFallback: (callback) => {
      scheduledReveal = callback;
      return { scheduled: true } as never;
    },
    clearRevealFallback: () => {
      scheduledReveal = null;
    },
  });

  window.loading = true;
  window.url = '';
  window.contentReady = false;

  const sent = runtime.sendToActiveOverlayWindow('jimaku:open', undefined, {
    restoreOnModalClose: 'jimaku',
  });

  assert.equal(sent, true);
  if (scheduledReveal === null) {
    throw new Error('expected reveal callback');
  }
  const runScheduledReveal: () => void = scheduledReveal;
  runScheduledReveal();

  assert.equal(window.getShowCount(), 0);

  runtime.notifyOverlayModalOpened('jimaku');
  assert.equal(window.getShowCount(), 1);
  assert.equal(window.ignoreMouseEvents, false);
});

test('sendToActiveOverlayWindow waits for modal ready-to-show before delivering open event', () => {
  const window = createMockWindow();
  window.contentReady = false;
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => null,
    getModalWindow: () => window as never,
    createModalWindow: () => {
      throw new Error('modal window should not be created when already present');
    },
    getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
    setModalWindowBounds: () => {},
  });

  const sent = runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });

  assert.equal(sent, true);
  assert.deepEqual(window.sent, []);
  window.emitDidFinishLoad();
  assert.deepEqual(window.sent, []);

  window.contentReady = true;
  window.emitReadyToShow();
  assert.deepEqual(window.sent, [['runtime-options:open']]);
});

test('sendToActiveOverlayWindow flushes every queued load and ready listener before sending', () => {
  const window = createMockWindow();
  window.contentReady = false;
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => null,
    getModalWindow: () => window as never,
    createModalWindow: () => {
      throw new Error('modal window should not be created when already present');
    },
    getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
    setModalWindowBounds: () => {},
  });

  assert.equal(
    runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
      restoreOnModalClose: 'runtime-options',
    }),
    true,
  );
  assert.equal(
    runtime.sendToActiveOverlayWindow('session-help:open', undefined, {
      restoreOnModalClose: 'session-help',
    }),
    true,
  );
  assert.deepEqual(window.sent, []);

  window.emitDidFinishLoad();
  assert.deepEqual(window.sent, []);

  window.contentReady = true;
  window.emitReadyToShow();
  assert.deepEqual(window.sent, [['runtime-options:open'], ['session-help:open']]);
});

test('modal reopen creates a fresh window after close destroys the previous one', () => {
  const firstWindow = createMockWindow();
  const secondWindow = createMockWindow();
  let currentModal: ReturnType<typeof createMockWindow> | null = firstWindow;

  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => null,
    getModalWindow: () => currentModal as never,
    createModalWindow: () => {
      currentModal = secondWindow;
      return secondWindow as never;
    },
    getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
    setModalWindowBounds: () => {},
  });

  runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });
  runtime.notifyOverlayModalOpened('runtime-options');
  runtime.handleOverlayModalClosed('runtime-options');

  assert.equal(firstWindow.isDestroyed(), true);

  const sent = runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });

  assert.equal(sent, true);
  assert.equal(currentModal, secondWindow);
  assert.equal(secondWindow.getShowCount(), 0);
});

test('modal reopen after close-destroy notifies state change on fresh window lifecycle', () => {
  const firstWindow = createMockWindow();
  const secondWindow = createMockWindow();
  let currentModal: ReturnType<typeof createMockWindow> | null = firstWindow;
  const state: boolean[] = [];

  const runtime = createOverlayModalRuntimeService(
    {
      getMainWindow: () => null,
      getModalWindow: () => currentModal as never,
      createModalWindow: () => {
        currentModal = secondWindow;
        return secondWindow as never;
      },
      getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
      setModalWindowBounds: () => {},
    },
    {
      onModalStateChange: (active: boolean): void => {
        state.push(active);
      },
    },
  );

  runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });
  runtime.notifyOverlayModalOpened('runtime-options');
  runtime.handleOverlayModalClosed('runtime-options');

  assert.deepEqual(state, [true, false]);
  assert.equal(firstWindow.isDestroyed(), true);

  runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });
  runtime.notifyOverlayModalOpened('runtime-options');

  assert.deepEqual(state, [true, false, true]);
  assert.equal(currentModal, secondWindow);
});

test('visible stale modal window is made interactive again before reopening', () => {
  const window = createMockWindow();
  window.visible = true;
  window.focused = true;
  window.webContentsFocused = false;
  window.ignoreMouseEvents = true;

  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => null,
    getModalWindow: () => window as never,
    createModalWindow: () => window as never,
    getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
    setModalWindowBounds: () => {},
  });

  const sent = runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });

  assert.equal(sent, true);
  assert.equal(window.ignoreMouseEvents, false);
  assert.equal(window.isFocused(), true);
  assert.equal(window.webContentsFocused, true);
  assert.deepEqual(window.sent, [['runtime-options:open']]);
});

test('waitForModalOpen resolves true after modal acknowledgement', async () => {
  const modalWindow = createMockWindow();
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => null,
    getModalWindow: () => modalWindow as never,
    createModalWindow: () => modalWindow as never,
    getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
    setModalWindowBounds: () => {},
  });

  runtime.sendToActiveOverlayWindow(
    'youtube:picker-open',
    { sessionId: 'yt-1' },
    {
      restoreOnModalClose: 'youtube-track-picker',
    },
  );
  const pending = runtime.waitForModalOpen('youtube-track-picker', 1000);
  runtime.notifyOverlayModalOpened('youtube-track-picker');

  assert.equal(await pending, true);
});

test('waitForModalOpen resolves false on timeout', async () => {
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => null,
    getModalWindow: () => null,
    createModalWindow: () => null,
    getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
    setModalWindowBounds: () => {},
  });

  assert.equal(await runtime.waitForModalOpen('youtube-track-picker', 5), false);
});

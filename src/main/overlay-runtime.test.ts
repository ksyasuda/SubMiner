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
  alwaysOnTopCalls: string[];
  showCount: number;
  hideCount: number;
  sent: unknown[][];
  loading: boolean;
  url: string;
  contentReady: boolean;
  documentLoaded: boolean;
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
  showInactive: () => void;
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
    alwaysOnTopCalls: [],
    showCount: 0,
    hideCount: 0,
    sent: [],
    loading: false,
    url: 'file:///overlay/index.html?layer=modal',
    contentReady: true,
    documentLoaded: true,
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
    setAlwaysOnTop: (flag: boolean, level?: string, relativeLevel?: number) => {
      state.alwaysOnTopCalls.push(`top:${flag}:${level ?? ''}:${relativeLevel ?? ''}`);
    },
    moveTop: () => {},
    getShowCount: () => state.showCount,
    getHideCount: () => state.hideCount,
    show: () => {
      state.visible = true;
      state.showCount += 1;
    },
    showInactive: () => {
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
      state.documentLoaded = true;
      (
        window as typeof window & { __subminerOverlayDocumentLoaded?: boolean }
      ).__subminerOverlayDocumentLoaded = true;
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

  Object.defineProperty(window, 'alwaysOnTopCalls', {
    get: () => state.alwaysOnTopCalls,
    set: (value: string[]) => {
      state.alwaysOnTopCalls = value;
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
      (
        window as typeof window & { __subminerOverlayContentReady?: boolean }
      ).__subminerOverlayContentReady = value;
    },
  });

  Object.defineProperty(window, 'documentLoaded', {
    get: () => state.documentLoaded,
    set: (value: boolean) => {
      state.documentLoaded = value;
      (
        window as typeof window & { __subminerOverlayDocumentLoaded?: boolean }
      ).__subminerOverlayDocumentLoaded = value;
    },
  });

  (
    window as typeof window & { __subminerOverlayContentReady?: boolean }
  ).__subminerOverlayContentReady = state.contentReady;
  (
    window as typeof window & { __subminerOverlayDocumentLoaded?: boolean }
  ).__subminerOverlayDocumentLoaded = state.documentLoaded;

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
  assert.deepEqual(calls, ['bounds:10,20,300,200', 'bounds:10,20,300,200']);
  assert.deepEqual(window.alwaysOnTopCalls, ['top:true:screen-saver:3']);
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

for (const platform of ['darwin', 'win32'] as const) {
  test(`primeModalWindow creates and warms a hidden modal on ${platform}`, () => {
    const modalWindow = createMockWindow();
    modalWindow.loading = true;
    modalWindow.url = '';
    modalWindow.contentReady = false;
    modalWindow.documentLoaded = false;
    let currentModal: ReturnType<typeof createMockWindow> | null = null;
    let createCalls = 0;
    const runtime = createOverlayModalRuntimeService(
      {
        getMainWindow: () => null,
        getModalWindow: () => currentModal as never,
        createModalWindow: () => {
          createCalls += 1;
          currentModal = modalWindow;
          return modalWindow as never;
        },
        getModalGeometry: () => ({ x: 1, y: 2, width: 300, height: 200 }),
        setModalWindowBounds: () => {},
      },
      { platform },
    );

    assert.equal(runtime.primeModalWindow(), true);
    assert.equal(createCalls, 1);
    assert.equal(modalWindow.isVisible(), false);

    modalWindow.loading = false;
    modalWindow.url = 'file:///overlay/index.html?layer=modal';
    modalWindow.emitDidFinishLoad();
    modalWindow.emitReadyToShow();
    modalWindow.contentReady = true;

    assert.equal(
      runtime.sendToActiveOverlayWindow('session-help:open', undefined, {
        restoreOnModalClose: 'session-help',
        preferModalWindow: true,
      }),
      true,
    );
    assert.equal(createCalls, 1);
    assert.equal(modalWindow.isVisible(), true);
    assert.deepEqual(modalWindow.sent, [['session-help:open']]);
  });
}

test('primeModalWindow leaves Linux modal creation lazy', () => {
  let createCalls = 0;
  const runtime = createOverlayModalRuntimeService(
    {
      getMainWindow: () => null,
      getModalWindow: () => null,
      createModalWindow: () => {
        createCalls += 1;
        return createMockWindow() as never;
      },
      getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
      setModalWindowBounds: () => {},
    },
    { platform: 'linux' },
  );

  assert.equal(runtime.primeModalWindow(), false);
  assert.equal(createCalls, 0);
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
  assert.deepEqual(window.sent, [['runtime-options:open']]);

  window.contentReady = true;
  window.emitReadyToShow();

  runtime.notifyOverlayModalOpened('runtime-options');
  assert.deepEqual(window.sent, [['runtime-options:open']]);
  assert.equal(window.getShowCount(), 1);
});

test('handleOverlayModalClosed keeps the modal window warm after all pending modals close', () => {
  const window = createMockWindow();
  const runtime = createOverlayModalRuntimeService(
    {
      getMainWindow: () => null,
      getModalWindow: () => window as never,
      createModalWindow: () => window as never,
      getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
      setModalWindowBounds: () => {},
    },
    { platform: 'darwin' },
  );

  runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });
  runtime.sendToActiveOverlayWindow(
    'subsync:open-manual',
    {
      ffsubsyncAvailable: true,
      videoReferenceAvailable: true,
      subtitleTracks: [],
      defaultReferenceTrackId: null,
      defaultTargetTrackId: null,
    },
    {
      restoreOnModalClose: 'subsync',
    },
  );

  runtime.handleOverlayModalClosed('runtime-options');
  assert.equal(window.isDestroyed(), false);

  runtime.handleOverlayModalClosed('subsync');
  assert.equal(window.isDestroyed(), false);
  assert.equal(window.isVisible(), false);
  assert.equal(window.ignoreMouseEvents, true);
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

test('modal window path restores visible main overlay before modal input deactivates', () => {
  const mainWindow = createMockWindow();
  mainWindow.visible = true;
  const modalWindow = createMockWindow();
  const events: string[] = [];
  const runtime = createOverlayModalRuntimeService(
    {
      getMainWindow: () => mainWindow as never,
      getModalWindow: () => modalWindow as never,
      createModalWindow: () => modalWindow as never,
      getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
      setModalWindowBounds: () => {},
    },
    {
      onModalStateChange: (active: boolean): void => {
        events.push(`state:${active}:visible:${mainWindow.isVisible()}`);
      },
    },
  );

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

  assert.equal(mainWindow.getShowCount(), 1);
  assert.equal(mainWindow.isVisible(), true);
  assert.deepEqual(events, ['state:true:visible:true', 'state:false:visible:true']);
});

test('macOS maps a new modal panel before focusing SubMiner and hiding the subtitle overlay', () => {
  const mainWindow = createMockWindow();
  mainWindow.visible = true;
  const modalWindow = createMockWindow();
  const events: string[] = [];
  const showInactive = modalWindow.showInactive;
  modalWindow.showInactive = () => {
    events.push('show-inactive');
    showInactive();
  };
  const hideMainWindow = mainWindow.hide;
  mainWindow.hide = () => {
    events.push('hide-main');
    hideMainWindow();
  };
  const runtime = createOverlayModalRuntimeService(
    {
      getMainWindow: () => mainWindow as never,
      getModalWindow: () => modalWindow as never,
      createModalWindow: () => modalWindow as never,
      getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
      setModalWindowBounds: () => {},
    },
    {
      platform: 'darwin',
      focusApplication: () => events.push('focus-application'),
    },
  );

  runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
    preferModalWindow: true,
  });
  runtime.notifyOverlayModalOpened('runtime-options');

  assert.deepEqual(events, ['show-inactive', 'focus-application', 'hide-main']);
  assert.equal(modalWindow.isVisible(), true);
  assert.equal(mainWindow.isVisible(), false);
});

test('modal window path runs final close handoff before modal input deactivates', () => {
  const mainWindow = createMockWindow();
  mainWindow.visible = true;
  const modalWindow = createMockWindow();
  const events: string[] = [];
  const runtime = createOverlayModalRuntimeService(
    {
      getMainWindow: () => mainWindow as never,
      getModalWindow: () => modalWindow as never,
      createModalWindow: () => modalWindow as never,
      getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
      setModalWindowBounds: () => {},
    },
    {
      onFinalModalClosed: (): void => {
        events.push(`handoff:visible:${mainWindow.isVisible()}`);
      },
      onModalStateChange: (active: boolean): void => {
        events.push(`state:${active}:visible:${mainWindow.isVisible()}`);
      },
    },
  );

  runtime.sendToActiveOverlayWindow(
    'youtube:picker-open',
    { sessionId: 'yt-1' },
    {
      restoreOnModalClose: 'youtube-track-picker',
      preferModalWindow: true,
    },
  );
  runtime.notifyOverlayModalOpened('youtube-track-picker');
  runtime.handleOverlayModalClosed('youtube-track-picker');

  assert.deepEqual(events, [
    'state:true:visible:true',
    'handoff:visible:true',
    'state:false:visible:true',
  ]);
});

test('modal runtime deactivates modal state when final close handoff throws', () => {
  const mainWindow = createMockWindow();
  mainWindow.visible = true;
  const modalWindow = createMockWindow();
  const events: string[] = [];
  const runtime = createOverlayModalRuntimeService(
    {
      getMainWindow: () => mainWindow as never,
      getModalWindow: () => modalWindow as never,
      createModalWindow: () => modalWindow as never,
      getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
      setModalWindowBounds: () => {},
    },
    {
      onFinalModalClosed: (): void => {
        events.push('handoff');
        throw new Error('handoff failed');
      },
      onModalStateChange: (active: boolean): void => {
        events.push(`state:${active}`);
      },
    },
  );

  runtime.sendToActiveOverlayWindow(
    'youtube:picker-open',
    { sessionId: 'yt-1' },
    {
      restoreOnModalClose: 'youtube-track-picker',
      preferModalWindow: true,
    },
  );
  runtime.notifyOverlayModalOpened('youtube-track-picker');

  assert.doesNotThrow(() => runtime.handleOverlayModalClosed('youtube-track-picker'));
  assert.deepEqual(events, ['state:true', 'handoff', 'state:false']);
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
    {
      ffsubsyncAvailable: true,
      videoReferenceAvailable: true,
      subtitleTracks: [],
      defaultReferenceTrackId: null,
      defaultTargetTrackId: null,
    },
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

test('handleOverlayModalClosed hides and retains modal window for single kiku modal', () => {
  const window = createMockWindow();
  const runtime = createOverlayModalRuntimeService(
    {
      getMainWindow: () => null,
      getModalWindow: () => window as never,
      createModalWindow: () => window as never,
      getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
      setModalWindowBounds: () => {},
    },
    { platform: 'darwin' },
  );

  runtime.sendToActiveOverlayWindow(
    'kiku:field-grouping-open',
    { test: true },
    {
      restoreOnModalClose: 'kiku',
    },
  );
  runtime.handleOverlayModalClosed('kiku');

  assert.equal(window.isDestroyed(), false);
  assert.equal(window.isVisible(), false);
  assert.equal(window.ignoreMouseEvents, true);
  assert.equal(runtime.getRestoreVisibleOverlayOnModalClose().size, 0);
});

test('modal fallback reveal skips showing window when content is not ready', async () => {
  const window = createMockWindow();
  let scheduledReveal: (() => void) | null = null;
  const runtime = createOverlayModalRuntimeService(
    {
      getMainWindow: () => null,
      getModalWindow: () => window as never,
      createModalWindow: () => {
        throw new Error('modal window should not be created when already present');
      },
      getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
      setModalWindowBounds: () => {},
    },
    {
      scheduleRevealFallback: (callback) => {
        scheduledReveal = callback;
        return { scheduled: true } as never;
      },
      clearRevealFallback: () => {
        scheduledReveal = null;
      },
    },
  );

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

test('sendToActiveOverlayWindow delivers on first modal load without waiting for ready-to-show', () => {
  const window = createMockWindow();
  window.loading = true;
  window.url = '';
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
  window.loading = false;
  window.url = 'file:///overlay/index.html?layer=modal';
  window.emitDidFinishLoad();
  assert.deepEqual(window.sent, [['runtime-options:open']]);

  window.contentReady = true;
  window.emitReadyToShow();
  assert.deepEqual(window.sent, [['runtime-options:open']]);
});

test('sendToActiveOverlayWindow delivers when the modal loaded before listeners were registered', () => {
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
  assert.deepEqual(window.sent, [['runtime-options:open']]);

  window.contentReady = true;
  window.emitReadyToShow();
  assert.deepEqual(window.sent, [['runtime-options:open']]);
});

test('sendToActiveOverlayWindow does not infer document readiness from a pending file URL', () => {
  const window = createMockWindow();
  window.contentReady = false;
  window.documentLoaded = false;
  window.loading = false;
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
  assert.deepEqual(window.sent, []);

  window.emitDidFinishLoad();
  assert.deepEqual(window.sent, [['runtime-options:open']]);
});

test('sendToActiveOverlayWindow rejects stale content readiness during document reload', () => {
  const window = createMockWindow();
  window.contentReady = true;
  window.documentLoaded = false;
  window.loading = false;
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
    runtime.sendToActiveOverlayWindow('session-help:open', undefined, {
      restoreOnModalClose: 'session-help',
    }),
    true,
  );
  assert.deepEqual(window.sent, []);

  window.emitDidFinishLoad();
  assert.deepEqual(window.sent, [['session-help:open']]);
});

test('sendToActiveOverlayWindow flushes every queued load and ready listener before sending', () => {
  const window = createMockWindow();
  window.loading = true;
  window.url = '';
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

  window.loading = false;
  window.url = 'file:///overlay/index.html?layer=modal';
  window.emitDidFinishLoad();
  assert.deepEqual(window.sent, [['runtime-options:open'], ['session-help:open']]);

  window.contentReady = true;
  window.emitReadyToShow();
  assert.deepEqual(window.sent, [['runtime-options:open'], ['session-help:open']]);
});

for (const platform of ['darwin', 'win32'] as const) {
  test(`modal reopen reuses the warm window and shows it immediately on ${platform}`, () => {
    const modalWindow = createMockWindow();
    let createCalls = 0;

    const runtime = createOverlayModalRuntimeService(
      {
        getMainWindow: () => null,
        getModalWindow: () => modalWindow as never,
        createModalWindow: () => {
          createCalls += 1;
          return modalWindow as never;
        },
        getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
        setModalWindowBounds: () => {},
      },
      { platform },
    );

    runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
      restoreOnModalClose: 'runtime-options',
    });
    runtime.notifyOverlayModalOpened('runtime-options');
    runtime.handleOverlayModalClosed('runtime-options');

    assert.equal(modalWindow.isDestroyed(), false);
    assert.equal(modalWindow.isVisible(), false);

    const sent = runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
      restoreOnModalClose: 'runtime-options',
    });

    assert.equal(sent, true);
    assert.equal(createCalls, 0);
    assert.equal(modalWindow.isVisible(), true);
    assert.equal(modalWindow.getShowCount(), 2);
  });
}

test('modal reopen on the warm window notifies state change for each lifecycle', () => {
  const modalWindow = createMockWindow();
  const state: boolean[] = [];

  const runtime = createOverlayModalRuntimeService(
    {
      getMainWindow: () => null,
      getModalWindow: () => modalWindow as never,
      createModalWindow: () => modalWindow as never,
      getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
      setModalWindowBounds: () => {},
    },
    {
      onModalStateChange: (active: boolean): void => {
        state.push(active);
      },
      platform: 'darwin',
    },
  );

  runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });
  runtime.notifyOverlayModalOpened('runtime-options');
  runtime.handleOverlayModalClosed('runtime-options');

  assert.deepEqual(state, [true, false]);
  assert.equal(modalWindow.isDestroyed(), false);

  runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });
  runtime.notifyOverlayModalOpened('runtime-options');

  assert.deepEqual(state, [true, false, true]);
  assert.equal(modalWindow.isVisible(), true);
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

test('waitForModalOpen resolves true when modal acknowledgement arrives before waiter registration', async () => {
  const modalWindow = createMockWindow();
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => null,
    getModalWindow: () => modalWindow as never,
    createModalWindow: () => modalWindow as never,
    getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
    setModalWindowBounds: () => {},
  });

  runtime.sendToActiveOverlayWindow(
    'kiku:field-grouping-request',
    {},
    {
      restoreOnModalClose: 'kiku',
    },
  );
  runtime.notifyOverlayModalOpened('kiku');

  assert.equal(await runtime.waitForModalOpen('kiku', 5), true);
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

test('modal placement reconcile retries until the Hyprland client is mapped', () => {
  const window = createMockWindow();
  const timers: Array<() => void> = [];
  const originalSetTimeout = globalThis.setTimeout;
  globalThis.setTimeout = ((cb: () => void) => {
    timers.push(cb);
    return { unref() {} };
  }) as unknown as typeof globalThis.setTimeout;

  const statuses: Array<{ applicable: boolean; clientFound: boolean }> = [];
  try {
    const runtime = createOverlayModalRuntimeService({
      getMainWindow: () => null,
      getModalWindow: () => window as never,
      createModalWindow: () => window as never,
      getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
      setModalWindowBounds: () => {
        // The compositor never maps the window, so every reconcile reports pending.
        const status = { applicable: true, clientFound: false, dispatched: false };
        statuses.push(status);
        return status;
      },
    });

    runtime.sendToActiveOverlayWindow(
      'kiku:field-grouping-open',
      { test: true },
      { restoreOnModalClose: 'kiku', preferModalWindow: true },
    );
    runtime.notifyOverlayModalOpened('kiku');

    let iterations = 0;
    while (timers.length > 0 && iterations < 50) {
      const next = timers.shift();
      next?.();
      iterations += 1;
    }

    // The reconcile ladder re-asserts placement across all six delays while the client
    // stays unmapped, instead of the old single post-show attempt.
    assert.ok(
      statuses.length >= 6,
      `expected at least 6 pending reconcile attempts, saw ${statuses.length}`,
    );
  } finally {
    globalThis.setTimeout = originalSetTimeout;
  }
});

test('modal placement reconcile stops retrying once the client is mapped', () => {
  const window = createMockWindow();
  const timers: Array<() => void> = [];
  const originalSetTimeout = globalThis.setTimeout;
  globalThis.setTimeout = ((cb: () => void) => {
    timers.push(cb);
    return { unref() {} };
  }) as unknown as typeof globalThis.setTimeout;

  let reconcileCount = 0;
  try {
    const runtime = createOverlayModalRuntimeService({
      getMainWindow: () => null,
      getModalWindow: () => window as never,
      createModalWindow: () => window as never,
      getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
      setModalWindowBounds: () => {
        reconcileCount += 1;
        // Client is already mapped, so placement is settled on the first attempt.
        return { applicable: true, clientFound: true, dispatched: true };
      },
    });

    runtime.sendToActiveOverlayWindow(
      'kiku:field-grouping-open',
      { test: true },
      { restoreOnModalClose: 'kiku', preferModalWindow: true },
    );
    runtime.notifyOverlayModalOpened('kiku');

    let iterations = 0;
    while (timers.length > 0 && iterations < 50) {
      const next = timers.shift();
      next?.();
      iterations += 1;
    }

    // No 6-deep ladder: a settled placement should not keep rescheduling.
    assert.ok(
      reconcileCount < 6,
      `expected the ladder to stop early, saw ${reconcileCount} reconcile attempts`,
    );
  } finally {
    globalThis.setTimeout = originalSetTimeout;
  }
});

test('modal placement reconcile cancels stale retry ladder after a newer visible modal interaction', () => {
  const window = createMockWindow();
  type TimerEntry = { active: boolean; callback: () => void };
  const timers: TimerEntry[] = [];
  const activeTimerCount = () => timers.filter((timer) => timer.active).length;
  const runNextActiveTimer = () => {
    const timer = timers.find((candidate) => candidate.active);
    if (!timer) return;
    timer.active = false;
    timer.callback();
  };
  const originalSetTimeout = globalThis.setTimeout;
  const originalClearTimeout = globalThis.clearTimeout;
  globalThis.setTimeout = ((cb: () => void) => {
    const timer = { active: true, callback: cb, unref() {} };
    timers.push(timer);
    return timer;
  }) as unknown as typeof globalThis.setTimeout;
  globalThis.clearTimeout = ((timeout: TimerEntry | undefined) => {
    if (timeout) {
      timeout.active = false;
    }
  }) as unknown as typeof globalThis.clearTimeout;

  try {
    const runtime = createOverlayModalRuntimeService({
      getMainWindow: () => null,
      getModalWindow: () => window as never,
      createModalWindow: () => window as never,
      getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
      setModalWindowBounds: () => ({ applicable: true, clientFound: false, dispatched: false }),
    });

    runtime.sendToActiveOverlayWindow(
      'kiku:field-grouping-open',
      { test: true },
      { restoreOnModalClose: 'kiku', preferModalWindow: true },
    );
    runtime.notifyOverlayModalOpened('kiku');
    assert.equal(activeTimerCount(), 1);

    runtime.sendToActiveOverlayWindow(
      'kiku:field-grouping-open',
      { test: true },
      { restoreOnModalClose: 'kiku', preferModalWindow: true },
    );
    assert.equal(activeTimerCount(), 2);

    runNextActiveTimer();

    assert.equal(activeTimerCount(), 1, 'stale retry should not schedule a continuation');
  } finally {
    globalThis.setTimeout = originalSetTimeout;
    globalThis.clearTimeout = originalClearTimeout;
  }
});

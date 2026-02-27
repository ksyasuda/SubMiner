import assert from 'node:assert/strict';
import test from 'node:test';
import { createOverlayModalRuntimeService } from './overlay-runtime';

type MockWindow = {
  destroyed: boolean;
  visible: boolean;
  focused: boolean;
  ignoreMouseEvents: boolean;
  webContentsFocused: boolean;
  showCount: number;
  hideCount: number;
  sent: unknown[][];
  loading: boolean;
  url: string;
  loadCallbacks: Array<() => void>;
};

function createMockWindow(): MockWindow & {
  isDestroyed: () => boolean;
  isVisible: () => boolean;
  isFocused: () => boolean;
  getURL: () => string;
  setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => void;
  getShowCount: () => number;
  getHideCount: () => number;
  show: () => void;
  hide: () => void;
  focus: () => void;
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
    webContentsFocused: false,
    showCount: 0,
    hideCount: 0,
    sent: [],
    loading: false,
    url: 'file:///overlay/index.html?layer=modal',
    loadCallbacks: [],
  };
  const window = {
    ...state,
    isDestroyed: () => state.destroyed,
    isVisible: () => state.visible,
    isFocused: () => state.focused,
    getURL: () => state.url,
    setIgnoreMouseEvents: (ignore: boolean, _options?: { forward?: boolean }) => {
      state.ignoreMouseEvents = ignore;
    },
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
    focus: () => {
      state.focused = true;
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

test('sendToActiveOverlayWindow waits for blank modal URL before sending open command', () => {
  const window = createMockWindow();
  window.url = '';
  window.loading = true;
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

  assert.equal(window.loadCallbacks.length, 1);
  window.loading = false;
  window.url = 'file:///overlay/index.html?layer=modal';
  window.loadCallbacks[0]!();

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
  runtime.sendToActiveOverlayWindow('subsync:open-manual', { sourceTracks: [] }, {
    restoreOnModalClose: 'subsync',
  });

  runtime.handleOverlayModalClosed('runtime-options');
  assert.equal(window.getHideCount(), 0);

  runtime.handleOverlayModalClosed('subsync');
  assert.equal(window.getHideCount(), 1);
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
  runtime.sendToActiveOverlayWindow('subsync:open-manual', { sourceTracks: [] }, {
    restoreOnModalClose: 'subsync',
  });
  assert.deepEqual(state, []);
  runtime.notifyOverlayModalOpened('runtime-options');
  assert.deepEqual(state, [true]);

  runtime.handleOverlayModalClosed('runtime-options');
  assert.deepEqual(state, [true]);

  runtime.handleOverlayModalClosed('subsync');
  assert.deepEqual(state, [true, false]);
});

test('handleOverlayModalClosed resets modal state even when modal window does not exist', () => {
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

  runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });
  runtime.notifyOverlayModalOpened('runtime-options');
  runtime.handleOverlayModalClosed('runtime-options');

  assert.deepEqual(state, [true, false]);
});

test('handleOverlayModalClosed hides modal window for single kiku modal', () => {
  const window = createMockWindow();
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => null,
    getModalWindow: () => window as never,
    createModalWindow: () => window as never,
    getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
    setModalWindowBounds: () => {},
  });

  runtime.sendToActiveOverlayWindow('kiku:field-grouping-open', { test: true }, {
    restoreOnModalClose: 'kiku',
  });
  runtime.handleOverlayModalClosed('kiku');

  assert.equal(window.getHideCount(), 1);
  assert.equal(runtime.getRestoreVisibleOverlayOnModalClose().size, 0);
});

test('modal fallback reveal keeps mouse events ignored until modal confirms open', async () => {
  const window = createMockWindow();
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => null,
    getModalWindow: () => window as never,
    createModalWindow: () => {
      throw new Error('modal window should not be created when already present');
    },
    getModalGeometry: () => ({ x: 0, y: 0, width: 400, height: 300 }),
    setModalWindowBounds: () => {},
  });

  window.loading = true;
  window.url = '';

  const sent = runtime.sendToActiveOverlayWindow('jimaku:open', undefined, {
    restoreOnModalClose: 'jimaku',
  });

  assert.equal(sent, true);
  assert.equal(window.ignoreMouseEvents, false);

  await new Promise<void>((resolve) => {
    setTimeout(resolve, 260);
  });

  assert.equal(window.getShowCount(), 1);
  assert.equal(window.ignoreMouseEvents, true);

  runtime.notifyOverlayModalOpened('jimaku');
  assert.equal(window.ignoreMouseEvents, false);
});

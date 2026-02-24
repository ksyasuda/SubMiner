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
  loadCallbacks: Array<() => void>;
};

function createMockWindow(): MockWindow & {
  isDestroyed: () => boolean;
  isVisible: () => boolean;
  isFocused: () => boolean;
  setIgnoreMouseEvents: (ignore: boolean) => void;
  getShowCount: () => number;
  getHideCount: () => number;
  show: () => void;
  hide: () => void;
  focus: () => void;
  webContents: {
    focused: boolean;
    isLoading: () => boolean;
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
    loadCallbacks: [],
  };
  return {
    ...state,
    isDestroyed: () => state.destroyed,
    isVisible: () => state.visible,
    isFocused: () => state.focused,
    setIgnoreMouseEvents: (ignore: boolean) => {
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
      send: (channel, payload) => {
        if (payload === undefined) {
          state.sent.push([channel]);
          return;
        }
        state.sent.push([channel, payload]);
      },
      focused: false,
      isFocused: () => state.webContentsFocused,
      once: (_event, cb) => {
        state.loadCallbacks.push(cb);
      },
      focus: () => {
        state.webContentsFocused = true;
      },
    },
  };
}

test('sendToActiveOverlayWindow targets modal window with full geometry and tracks close restore', () => {
  const window = createMockWindow();
  const calls: string[] = [];
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => null,
    getInvisibleWindow: () => null,
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
  assert.equal(window.getShowCount(), 1);
  assert.equal(window.isFocused(), true);
  assert.deepEqual(window.sent, [['runtime-options:open']]);
});

test('sendToActiveOverlayWindow creates modal window lazily when absent', () => {
  const window = createMockWindow();
  let modalWindow: ReturnType<typeof createMockWindow> | null = null;
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => null,
    getInvisibleWindow: () => null,
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
  assert.deepEqual(window.sent, [['jimaku:open']]);
});

test('handleOverlayModalClosed hides modal window only after all pending modals close', () => {
  const window = createMockWindow();
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => null,
    getInvisibleWindow: () => null,
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

test('modal runtime notifies callers when modal input state becomes active/inactive', () => {
  const window = createMockWindow();
  const state: boolean[] = [];
  const runtime = createOverlayModalRuntimeService(
    {
      getMainWindow: () => null,
      getInvisibleWindow: () => null,
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
  assert.deepEqual(state, [true]);

  runtime.handleOverlayModalClosed('runtime-options');
  assert.deepEqual(state, [true]);

  runtime.handleOverlayModalClosed('subsync');
  assert.deepEqual(state, [true, false]);
});

test('handleOverlayModalClosed hides modal window for single kiku modal', () => {
  const window = createMockWindow();
  const runtime = createOverlayModalRuntimeService({
    getMainWindow: () => null,
    getInvisibleWindow: () => null,
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

import assert from 'node:assert/strict';
import test from 'node:test';

import { createKeyboardHandlers } from './keyboard.js';
import { createRendererState } from '../state.js';
import { YOMITAN_POPUP_COMMAND_EVENT } from '../yomitan-popup.js';

type CommandEventDetail = {
  type?: string;
  visible?: boolean;
  key?: string;
  code?: string;
};

function createClassList() {
  const classes = new Set<string>();
  return {
    add: (...tokens: string[]) => {
      for (const token of tokens) {
        classes.add(token);
      }
    },
    remove: (...tokens: string[]) => {
      for (const token of tokens) {
        classes.delete(token);
      }
    },
    contains: (token: string) => classes.has(token),
  };
}

function wait(ms: number): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}

function installKeyboardTestGlobals() {
  const previousWindow = (globalThis as { window?: unknown }).window;
  const previousDocument = (globalThis as { document?: unknown }).document;
  const previousMutationObserver = (globalThis as { MutationObserver?: unknown }).MutationObserver;
  const previousCustomEvent = (globalThis as { CustomEvent?: unknown }).CustomEvent;
  const previousMouseEvent = (globalThis as { MouseEvent?: unknown }).MouseEvent;

  const documentListeners = new Map<string, Array<(event: unknown) => void>>();
  const commandEvents: CommandEventDetail[] = [];

  let popupVisible = false;

  const popupIframe = {
    tagName: 'IFRAME',
    classList: {
      contains: (token: string) => token === 'yomitan-popup',
    },
    id: 'yomitan-popup-1',
    getBoundingClientRect: () => ({ left: 0, top: 0, width: 100, height: 100 }),
  };

  const selection = {
    removeAllRanges: () => {},
    addRange: () => {},
  };

  const overlayFocusCalls: Array<{ preventScroll?: boolean }> = [];
  let focusMainWindowCalls = 0;
  let windowFocusCalls = 0;

  class TestCustomEvent extends Event {
    detail: unknown;

    constructor(type: string, init?: { detail?: unknown }) {
      super(type);
      this.detail = init?.detail;
    }
  }

  class TestMouseEvent extends Event {
    constructor(type: string) {
      super(type);
    }
  }

  Object.defineProperty(globalThis, 'CustomEvent', {
    configurable: true,
    value: TestCustomEvent,
  });

  Object.defineProperty(globalThis, 'MouseEvent', {
    configurable: true,
    value: TestMouseEvent,
  });

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      addEventListener: () => {},
      dispatchEvent: (event: Event) => {
        if (event.type === YOMITAN_POPUP_COMMAND_EVENT) {
          const detail = (event as Event & { detail?: CommandEventDetail }).detail;
          commandEvents.push(detail ?? {});
        }
        return true;
      },
      getComputedStyle: () => ({
        visibility: 'visible',
        display: 'block',
        opacity: '1',
      }),
      getSelection: () => selection,
      focus: () => {
        windowFocusCalls += 1;
      },
      electronAPI: {
        getKeybindings: async () => [],
        sendMpvCommand: () => {},
        toggleDevTools: () => {},
        focusMainWindow: () => {
          focusMainWindowCalls += 1;
          return Promise.resolve();
        },
      },
    },
  });

  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      addEventListener: (type: string, listener: (event: unknown) => void) => {
        const listeners = documentListeners.get(type) ?? [];
        listeners.push(listener);
        documentListeners.set(type, listeners);
      },
      querySelectorAll: () => {
        if (popupVisible) {
          return [popupIframe];
        }
        return [];
      },
      createRange: () => ({
        selectNodeContents: () => {},
      }),
      body: {},
    },
  });

  Object.defineProperty(globalThis, 'MutationObserver', {
    configurable: true,
    value: class {
      observe() {}
    },
  });

  function dispatchKeydown(event: {
    key: string;
    code: string;
    ctrlKey?: boolean;
    metaKey?: boolean;
    altKey?: boolean;
    shiftKey?: boolean;
    repeat?: boolean;
  }): void {
    const listeners = documentListeners.get('keydown') ?? [];
    const keyboardEvent = {
      key: event.key,
      code: event.code,
      ctrlKey: event.ctrlKey ?? false,
      metaKey: event.metaKey ?? false,
      altKey: event.altKey ?? false,
      shiftKey: event.shiftKey ?? false,
      repeat: event.repeat ?? false,
      preventDefault: () => {},
      target: null,
    };
    for (const listener of listeners) {
      listener(keyboardEvent);
    }
  }

  function dispatchFocusInOnPopup(): void {
    const listeners = documentListeners.get('focusin') ?? [];
    const focusEvent = {
      target: popupIframe,
    };
    for (const listener of listeners) {
      listener(focusEvent);
    }
  }

  function restore() {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
    Object.defineProperty(globalThis, 'MutationObserver', {
      configurable: true,
      value: previousMutationObserver,
    });
    Object.defineProperty(globalThis, 'CustomEvent', {
      configurable: true,
      value: previousCustomEvent,
    });
    Object.defineProperty(globalThis, 'MouseEvent', {
      configurable: true,
      value: previousMouseEvent,
    });
  }

  const overlay = {
    focus: (options?: { preventScroll?: boolean }) => {
      overlayFocusCalls.push(options ?? {});
    },
  };

  return {
    commandEvents,
    overlay,
    overlayFocusCalls,
    focusMainWindowCalls: () => focusMainWindowCalls,
    windowFocusCalls: () => windowFocusCalls,
    dispatchKeydown,
    dispatchFocusInOnPopup,
    setPopupVisible: (value: boolean) => {
      popupVisible = value;
    },
    restore,
  };
}

function createKeyboardHandlerHarness() {
  const testGlobals = installKeyboardTestGlobals();
  const subtitleRootClassList = createClassList();

  const wordNodes = [
    {
      classList: createClassList(),
      getBoundingClientRect: () => ({ left: 10, top: 10, width: 30, height: 20 }),
      dispatchEvent: () => true,
    },
    {
      classList: createClassList(),
      getBoundingClientRect: () => ({ left: 80, top: 10, width: 30, height: 20 }),
      dispatchEvent: () => true,
    },
    {
      classList: createClassList(),
      getBoundingClientRect: () => ({ left: 150, top: 10, width: 30, height: 20 }),
      dispatchEvent: () => true,
    },
  ];

  const ctx = {
    dom: {
      subtitleRoot: {
        classList: subtitleRootClassList,
        querySelectorAll: () => wordNodes,
      },
      subtitleContainer: {
        contains: () => false,
      },
      overlay: testGlobals.overlay,
    },
    platform: {
      shouldToggleMouseIgnore: false,
      isMacOSPlatform: false,
      overlayLayer: 'always-on-top',
    },
    state: createRendererState(),
  };

  const handlers = createKeyboardHandlers(ctx as never, {
    handleRuntimeOptionsKeydown: () => false,
    handleSubsyncKeydown: () => false,
    handleKikuKeydown: () => false,
    handleJimakuKeydown: () => false,
    handleSessionHelpKeydown: () => false,
    openSessionHelpModal: () => {},
    appendClipboardVideoToQueue: () => {},
  });

  return { ctx, handlers, testGlobals };
}

test('keyboard mode: left and right move token selection while popup remains open', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    ctx.state.keyboardSelectedWordIndex = 1;
    ctx.state.yomitanPopupVisible = true;
    testGlobals.setPopupVisible(true);

    testGlobals.dispatchKeydown({ key: 'ArrowRight', code: 'ArrowRight' });
    assert.equal(ctx.state.keyboardSelectedWordIndex, 2);

    testGlobals.dispatchKeydown({ key: 'ArrowLeft', code: 'ArrowLeft' });
    assert.equal(ctx.state.keyboardSelectedWordIndex, 1);
    await wait(0);

    const closeEvents = testGlobals.commandEvents.filter(
      (event) => event.type === 'setVisible' && event.visible === false,
    );
    assert.equal(closeEvents.length, 0);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: up and j open yomitan lookup for selected token', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    testGlobals.dispatchKeydown({ key: 'ArrowUp', code: 'ArrowUp' });
    testGlobals.dispatchKeydown({ key: 'j', code: 'KeyJ' });

    await wait(80);

    const openEvents = testGlobals.commandEvents.filter((event) => event.type === 'scanSelectedText');
    assert.equal(openEvents.length, 2);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: down closes yomitan lookup window', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    ctx.state.yomitanPopupVisible = true;
    testGlobals.setPopupVisible(true);

    testGlobals.dispatchKeydown({ key: 'ArrowDown', code: 'ArrowDown' });
    await wait(0);

    const closeEvents = testGlobals.commandEvents.filter(
      (event) => event.type === 'setVisible' && event.visible === false,
    );
    assert.equal(closeEvents.length, 1);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: h moves left when popup is closed', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    ctx.state.keyboardSelectedWordIndex = 2;
    ctx.state.yomitanPopupVisible = false;
    testGlobals.setPopupVisible(false);

    testGlobals.dispatchKeydown({ key: 'h', code: 'KeyH' });
    assert.equal(ctx.state.keyboardSelectedWordIndex, 1);

    const closeEvents = testGlobals.commandEvents.filter(
      (event) => event.type === 'setVisible' && event.visible === false,
    );
    assert.equal(closeEvents.length, 0);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: h moves left while popup is open and keeps lookup active', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    ctx.state.keyboardSelectedWordIndex = 2;
    ctx.state.yomitanPopupVisible = true;
    testGlobals.setPopupVisible(true);

    testGlobals.dispatchKeydown({ key: 'h', code: 'KeyH' });
    await wait(80);

    assert.equal(ctx.state.keyboardSelectedWordIndex, 1);
    const openEvents = testGlobals.commandEvents.filter((event) => event.type === 'scanSelectedText');
    assert.equal(openEvents.length > 0, true);
    const closeEvents = testGlobals.commandEvents.filter(
      (event) => event.type === 'setVisible' && event.visible === false,
    );
    assert.equal(closeEvents.length, 0);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: opening lookup restores overlay keyboard focus', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    testGlobals.dispatchKeydown({ key: 'ArrowUp', code: 'ArrowUp' });
    await wait(0);

    assert.equal(testGlobals.focusMainWindowCalls() > 0, true);
    assert.equal(testGlobals.windowFocusCalls() > 0, true);
    assert.equal(testGlobals.overlayFocusCalls.length > 0, true);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: popup iframe focusin reclaims overlay keyboard focus', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();
    testGlobals.setPopupVisible(true);

    const before = testGlobals.focusMainWindowCalls();
    testGlobals.dispatchFocusInOnPopup();
    await wait(260);

    assert.equal(testGlobals.focusMainWindowCalls() > before, true);
    assert.equal(testGlobals.windowFocusCalls() > 0, true);
    assert.equal(testGlobals.overlayFocusCalls.length > 0, true);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

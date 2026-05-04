import assert from 'node:assert/strict';
import test from 'node:test';

import { createKeyboardHandlers } from './keyboard.js';
import { createRendererState } from '../state.js';
import type { CompiledSessionBinding } from '../../types';
import { YOMITAN_POPUP_COMMAND_EVENT, YOMITAN_POPUP_HIDDEN_EVENT } from '../yomitan-popup.js';

type CommandEventDetail = {
  type?: string;
  visible?: boolean;
  key?: string;
  code?: string;
  repeat?: boolean;
  direction?: number;
  deltaX?: number;
  deltaY?: number;
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
  const windowListeners = new Map<string, Array<(event: unknown) => void>>();
  const commandEvents: CommandEventDetail[] = [];
  const mpvCommands: Array<Array<string | number>> = [];
  const sessionActions: Array<{ actionId: string; payload?: unknown }> = [];
  let sessionBindings: CompiledSessionBinding[] = [];
  let playbackPausedResponse: boolean | null = false;
  let statsToggleKey = 'Backquote';
  let markWatchedKey = 'KeyW';
  let configuredShortcuts = {
    copySubtitle: '',
    copySubtitleMultiple: '',
    updateLastCardFromClipboard: '',
    triggerFieldGrouping: '',
    triggerSubsync: 'Ctrl+Alt+S',
    mineSentence: '',
    mineSentenceMultiple: '',
    multiCopyTimeoutMs: 1000,
    toggleSecondarySub: '',
    markAudioCard: '',
    openRuntimeOptions: 'CommandOrControl+Shift+O',
    openJimaku: 'Ctrl+Shift+J',
    openSessionHelp: 'CommandOrControl+Shift+H',
    openControllerSelect: 'Alt+C',
    openControllerDebug: 'Alt+Shift+C',
    toggleSubtitleSidebar: '',
    toggleVisibleOverlayGlobal: '',
  };
  let markActiveVideoWatchedResult = true;
  let markActiveVideoWatchedCalls = 0;
  let statsToggleOverlayCalls = 0;
  const openedModalNotifications: string[] = [];
  let selectionClearCount = 0;
  let selectionAddCount = 0;

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
    removeAllRanges: () => {
      selectionClearCount += 1;
    },
    addRange: () => {
      selectionAddCount += 1;
    },
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
      addEventListener: (type: string, listener: (event: unknown) => void) => {
        const listeners = windowListeners.get(type) ?? [];
        listeners.push(listener);
        windowListeners.set(type, listeners);
      },
      dispatchEvent: (event: Event) => {
        if (event.type === YOMITAN_POPUP_COMMAND_EVENT) {
          const detail = (event as Event & { detail?: CommandEventDetail }).detail;
          commandEvents.push(detail ?? {});
        }
        const listeners = windowListeners.get(event.type) ?? [];
        for (const listener of listeners) {
          listener(event);
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
        getSessionBindings: async () => sessionBindings,
        getConfiguredShortcuts: async () => configuredShortcuts,
        sendMpvCommand: (command: Array<string | number>) => {
          mpvCommands.push(command);
        },
        dispatchSessionAction: async (actionId: string, payload?: unknown) => {
          sessionActions.push({ actionId, payload });
        },
        getPlaybackPaused: async () => playbackPausedResponse,
        getStatsToggleKey: async () => statsToggleKey,
        getMarkWatchedKey: async () => markWatchedKey,
        markActiveVideoWatched: async () => {
          markActiveVideoWatchedCalls += 1;
          return markActiveVideoWatchedResult;
        },
        toggleDevTools: () => {},
        toggleStatsOverlay: () => {
          statsToggleOverlayCalls += 1;
        },
        focusMainWindow: () => {
          focusMainWindowCalls += 1;
          return Promise.resolve();
        },
        notifyOverlayModalOpened: (modal: string) => {
          openedModalNotifications.push(modal);
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

  function dispatchWindowEvent(type: string): void {
    const listeners = windowListeners.get(type) ?? [];
    for (const listener of listeners) {
      listener(new Event(type));
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
    mpvCommands,
    sessionActions,
    overlay,
    overlayFocusCalls,
    focusMainWindowCalls: () => focusMainWindowCalls,
    windowFocusCalls: () => windowFocusCalls,
    dispatchKeydown,
    dispatchFocusInOnPopup,
    dispatchWindowEvent,
    setPopupVisible: (value: boolean) => {
      popupVisible = value;
    },
    setStatsToggleKey: (value: string) => {
      statsToggleKey = value;
    },
    setMarkWatchedKey: (value: string) => {
      markWatchedKey = value;
    },
    setConfiguredShortcuts: (value: typeof configuredShortcuts) => {
      configuredShortcuts = value;
    },
    setSessionBindings: (value: CompiledSessionBinding[]) => {
      sessionBindings = value;
    },
    setMarkActiveVideoWatchedResult: (value: boolean) => {
      markActiveVideoWatchedResult = value;
    },
    markActiveVideoWatchedCalls: () => markActiveVideoWatchedCalls,
    statsToggleOverlayCalls: () => statsToggleOverlayCalls,
    openedModalNotifications,
    getPlaybackPaused: async () => playbackPausedResponse,
    setPlaybackPausedResponse: (value: boolean | null) => {
      playbackPausedResponse = value;
    },
    selectionClearCount: () => selectionClearCount,
    selectionAddCount: () => selectionAddCount,
    restore,
  };
}

function createKeyboardHandlerHarness() {
  const testGlobals = installKeyboardTestGlobals();
  const subtitleRootClassList = createClassList();
  const subtitleContainerClassList = createClassList();
  let controllerSelectKeydownCount = 0;
  let openControllerSelectCount = 0;
  let openControllerDebugCount = 0;
  let playlistBrowserKeydownCount = 0;

  const createWordNode = (left: number) => ({
    classList: createClassList(),
    getBoundingClientRect: () => ({ left, top: 10, width: 30, height: 20 }),
    dispatchEvent: () => true,
  });
  let wordNodes = [createWordNode(10), createWordNode(80), createWordNode(150)];

  const ctx = {
    dom: {
      subtitleRoot: {
        classList: subtitleRootClassList,
        querySelectorAll: () => wordNodes,
      },
      subtitleContainer: {
        classList: subtitleContainerClassList,
        contains: () => false,
      },
      overlay: testGlobals.overlay,
    },
    platform: {
      isLinuxPlatform: false,
      shouldToggleMouseIgnore: false,
      isMacOSPlatform: false,
      isModalLayer: false,
      overlayLayer: 'always-on-top',
    },
    state: createRendererState(),
  };

  const handlers = createKeyboardHandlers(ctx as never, {
    handleRuntimeOptionsKeydown: () => false,
    handleCharacterDictionaryKeydown: () => false,
    handleSubsyncKeydown: () => false,
    handleKikuKeydown: () => false,
    handleJimakuKeydown: () => false,
    handleControllerSelectKeydown: () => {
      controllerSelectKeydownCount += 1;
      return true;
    },
    handleControllerDebugKeydown: () => false,
    handleYoutubePickerKeydown: () => false,
    handlePlaylistBrowserKeydown: () => {
      playlistBrowserKeydownCount += 1;
      return true;
    },
    handleSessionHelpKeydown: () => false,
    openSessionHelpModal: () => {},
    openControllerSelectModal: () => {
      openControllerSelectCount += 1;
    },
    openControllerDebugModal: () => {
      openControllerDebugCount += 1;
    },
    appendClipboardVideoToQueue: () => {},
    getPlaybackPaused: () => testGlobals.getPlaybackPaused(),
  });

  return {
    ctx,
    handlers,
    testGlobals,
    controllerSelectKeydownCount: () => controllerSelectKeydownCount,
    openControllerSelectCount: () => openControllerSelectCount,
    openControllerDebugCount: () => openControllerDebugCount,
    playlistBrowserKeydownCount: () => playlistBrowserKeydownCount,
    setWordCount: (count: number) => {
      wordNodes = Array.from({ length: count }, (_, index) => createWordNode(10 + index * 70));
    },
  };
}

test('primary subtitle visibility key hides and restores the subtitle bar without mpv sub-visibility', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();

    testGlobals.dispatchKeydown({ key: 'v', code: 'KeyV' });
    assert.equal(ctx.dom.subtitleContainer.classList.contains('primary-sub-hidden'), true);

    testGlobals.dispatchKeydown({ key: 'v', code: 'KeyV' });
    assert.equal(ctx.dom.subtitleContainer.classList.contains('primary-sub-hidden'), false);
    assert.equal(
      testGlobals.mpvCommands.some((command) => command.includes('sub-visibility')),
      false,
    );
  } finally {
    testGlobals.restore();
  }
});

test('session help chord resolver follows remapped session bindings', async () => {
  const { handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();

    assert.deepEqual(handlers.getSessionHelpOpeningInfo(), {
      bindingKey: 'KeyH',
      fallbackUsed: false,
      fallbackUnavailable: false,
    });

    handlers.updateSessionBindings([
      {
        sourcePath: 'keybindings[0].key',
        originalKey: 'KeyH',
        key: { code: 'KeyH', modifiers: [] },
        actionType: 'session-action',
        actionId: 'openJimaku',
      },
      {
        sourcePath: 'keybindings[1].key',
        originalKey: 'KeyJ',
        key: { code: 'KeyJ', modifiers: [] },
        actionType: 'mpv-command',
        command: ['cycle', 'pause'],
      },
    ] as never);

    assert.deepEqual(handlers.getSessionHelpOpeningInfo(), {
      bindingKey: 'KeyK',
      fallbackUsed: true,
      fallbackUnavailable: false,
    });

    handlers.updateSessionBindings([
      {
        sourcePath: 'keybindings[0].key',
        originalKey: 'KeyH',
        key: { code: 'KeyH', modifiers: [] },
        actionType: 'session-action',
        actionId: 'openSessionHelp',
      },
      {
        sourcePath: 'keybindings[1].key',
        originalKey: 'KeyK',
        key: { code: 'KeyK', modifiers: [] },
        actionType: 'session-action',
        actionId: 'openControllerSelect',
      },
    ] as never);

    assert.deepEqual(handlers.getSessionHelpOpeningInfo(), {
      bindingKey: 'KeyK',
      fallbackUsed: true,
      fallbackUnavailable: true,
    });
  } finally {
    testGlobals.restore();
  }
});

test('numeric selection ignores non-digit keys instead of falling through to other shortcuts', async () => {
  const { handlers, testGlobals, ctx } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.beginSessionNumericSelection('copySubtitleMultiple');

    testGlobals.dispatchKeydown({ key: 'y', code: 'KeyY' });

    assert.equal(ctx.state.chordPending, false);
    assert.deepEqual(testGlobals.sessionActions, []);
    assert.equal(
      testGlobals.commandEvents.some((event) => event.type === 'forwardKeyDown'),
      false,
    );
  } finally {
    testGlobals.restore();
  }
});

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

test('keyboard mode: up/down/j/k do not open or close lookup when popup is closed', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    testGlobals.dispatchKeydown({ key: 'ArrowUp', code: 'ArrowUp' });
    testGlobals.dispatchKeydown({ key: 'ArrowDown', code: 'ArrowDown' });
    testGlobals.dispatchKeydown({ key: 'j', code: 'KeyJ' });
    testGlobals.dispatchKeydown({ key: 'k', code: 'KeyK' });

    await wait(0);

    const openEvents = testGlobals.commandEvents.filter(
      (event) => event.type === 'scanSelectedText',
    );
    assert.equal(openEvents.length, 0);
    const closeEvents = testGlobals.commandEvents.filter(
      (event) => event.type === 'setVisible' && event.visible === false,
    );
    assert.equal(closeEvents.length, 0);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: up/down/j/k forward keydown to yomitan popup when open', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    ctx.state.yomitanPopupVisible = true;
    testGlobals.setPopupVisible(true);

    testGlobals.dispatchKeydown({ key: 'ArrowUp', code: 'ArrowUp' });
    testGlobals.dispatchKeydown({ key: 'ArrowDown', code: 'ArrowDown' });
    testGlobals.dispatchKeydown({ key: 'j', code: 'KeyJ' });
    testGlobals.dispatchKeydown({ key: 'k', code: 'KeyK' });

    const forwarded = testGlobals.commandEvents.filter((event) => event.type === 'forwardKeyDown');
    assert.equal(forwarded.length, 4);
    assert.equal(
      forwarded.some((event) => event.code === 'ArrowUp'),
      true,
    );
    assert.equal(
      forwarded.some((event) => event.code === 'ArrowDown'),
      true,
    );
    assert.equal(
      forwarded.some((event) => event.code === 'KeyJ'),
      true,
    );
    assert.equal(
      forwarded.some((event) => event.code === 'KeyK'),
      true,
    );

    const openEvents = testGlobals.commandEvents.filter(
      (event) => event.type === 'scanSelectedText',
    );
    assert.equal(openEvents.length, 0);
    const closeEvents = testGlobals.commandEvents.filter(
      (event) => event.type === 'setVisible' && event.visible === false,
    );
    assert.equal(closeEvents.length, 0);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: popup keybinds take precedence over configured session bindings', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.updateSessionBindings([
      {
        sourcePath: 'keybindings[0].key',
        originalKey: 'KeyJ',
        key: { code: 'KeyJ', modifiers: [] },
        actionType: 'mpv-command',
        command: ['cycle', 'sid'],
      },
    ] as never);
    handlers.handleKeyboardModeToggleRequested();

    ctx.state.yomitanPopupVisible = true;
    testGlobals.setPopupVisible(true);

    testGlobals.dispatchKeydown({ key: 'j', code: 'KeyJ' });

    assert.deepEqual(testGlobals.mpvCommands, []);
    assert.deepEqual(
      testGlobals.commandEvents
        .filter((event) => event.type === 'forwardKeyDown')
        .map((event) => event.code),
      ['KeyJ'],
    );
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: repeated popup navigation keys are forwarded while popup is open', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    ctx.state.yomitanPopupVisible = true;
    testGlobals.setPopupVisible(true);

    testGlobals.dispatchKeydown({ key: 'j', code: 'KeyJ', repeat: true });
    testGlobals.dispatchKeydown({ key: 'ArrowDown', code: 'ArrowDown', repeat: true });

    const forwarded = testGlobals.commandEvents.filter((event) => event.type === 'forwardKeyDown');
    assert.equal(forwarded.length, 2);
    assert.deepEqual(
      forwarded.map((event) => ({ code: event.code, repeat: event.repeat })),
      [
        { code: 'KeyJ', repeat: true },
        { code: 'ArrowDown', repeat: true },
      ],
    );
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('popup-visible mpv keybindings still fire for bound keys', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.updateSessionBindings([
      {
        sourcePath: 'keybindings[0].key',
        originalKey: 'Space',
        key: { code: 'Space', modifiers: [] },
        actionType: 'mpv-command',
        command: ['cycle', 'pause'],
      },
      {
        sourcePath: 'keybindings[1].key',
        originalKey: 'KeyQ',
        key: { code: 'KeyQ', modifiers: [] },
        actionType: 'mpv-command',
        command: ['quit'],
      },
    ] as never);

    ctx.state.yomitanPopupVisible = true;
    testGlobals.setPopupVisible(true);

    testGlobals.dispatchKeydown({ key: ' ', code: 'Space' });
    testGlobals.dispatchKeydown({ key: 'q', code: 'KeyQ' });

    assert.deepEqual(testGlobals.mpvCommands.slice(-2), [['cycle', 'pause'], ['quit']]);
  } finally {
    testGlobals.restore();
  }
});

test('paused configured subtitle-jump keybinding re-applies pause after backward seek', async () => {
  const { handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.updateSessionBindings([
      {
        sourcePath: 'keybindings[0].key',
        originalKey: 'Shift+KeyH',
        key: { code: 'KeyH', modifiers: ['shift'] },
        actionType: 'mpv-command',
        command: ['sub-seek', -1],
      },
    ] as never);
    testGlobals.setPlaybackPausedResponse(true);

    testGlobals.dispatchKeydown({ key: 'H', code: 'KeyH', shiftKey: true });
    await wait(0);

    assert.deepEqual(testGlobals.mpvCommands.slice(-2), [
      ['sub-seek', -1],
      ['set_property', 'pause', 'yes'],
    ]);
  } finally {
    testGlobals.restore();
  }
});

test('configured subtitle-jump keybinding preserves pause when pause state is unknown', async () => {
  const { handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.updateSessionBindings([
      {
        sourcePath: 'keybindings[0].key',
        originalKey: 'Shift+KeyH',
        key: { code: 'KeyH', modifiers: ['shift'] },
        actionType: 'mpv-command',
        command: ['sub-seek', -1],
      },
    ] as never);
    testGlobals.setPlaybackPausedResponse(null);

    testGlobals.dispatchKeydown({ key: 'H', code: 'KeyH', shiftKey: true });
    await wait(0);

    assert.deepEqual(testGlobals.mpvCommands.slice(-2), [
      ['sub-seek', -1],
      ['set_property', 'pause', 'yes'],
    ]);
  } finally {
    testGlobals.restore();
  }
});

test('visible-layer y-t dispatches mpv plugin toggle while overlay owns focus', async () => {
  const { handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();

    testGlobals.dispatchKeydown({ key: 'y', code: 'KeyY' });
    testGlobals.dispatchKeydown({ key: 't', code: 'KeyT' });

    assert.equal(
      testGlobals.mpvCommands.some(
        (command) => command[0] === 'script-message' && command[1] === 'subminer-toggle',
      ),
      true,
    );
  } finally {
    testGlobals.restore();
  }
});

test('refreshConfiguredShortcuts updates hot-reloaded stats and watched keys', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();

    testGlobals.setConfiguredShortcuts({
      copySubtitle: '',
      copySubtitleMultiple: '',
      updateLastCardFromClipboard: '',
      triggerFieldGrouping: '',
      triggerSubsync: 'Ctrl+Alt+S',
      mineSentence: '',
      mineSentenceMultiple: '',
      multiCopyTimeoutMs: 3333,
      toggleSecondarySub: '',
      markAudioCard: '',
      openRuntimeOptions: 'CommandOrControl+Shift+O',
      openJimaku: 'Ctrl+Shift+J',
      openSessionHelp: 'CommandOrControl+Shift+H',
      openControllerSelect: 'Alt+C',
      openControllerDebug: 'Alt+Shift+C',
      toggleSubtitleSidebar: '',
      toggleVisibleOverlayGlobal: '',
    });
    testGlobals.setStatsToggleKey('');
    testGlobals.setMarkWatchedKey('');

    await handlers.refreshConfiguredShortcuts();

    assert.equal(ctx.state.sessionActionTimeoutMs, 3333);
    assert.equal(ctx.state.statsToggleKey, '');
    assert.equal(ctx.state.markWatchedKey, '');
  } finally {
    testGlobals.restore();
  }
});

test('keyboard mode: controller helpers dispatch popup audio play/cycle and scroll bridge commands', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    ctx.state.yomitanPopupVisible = true;
    testGlobals.setPopupVisible(true);

    assert.equal(handlers.playCurrentAudioForController(), true);
    assert.equal(handlers.cyclePopupAudioSourceForController(1), true);
    assert.equal(handlers.scrollPopupByController(48, -24), true);

    assert.deepEqual(testGlobals.commandEvents.slice(-3), [
      { type: 'playCurrentAudio' },
      { type: 'cycleAudioSource', direction: 1 },
      { type: 'scrollBy', deltaX: 48, deltaY: -24 },
    ]);
  } finally {
    testGlobals.restore();
  }
});

test('keyboard mode: configured controller select binding opens locally without dispatching a session action', async () => {
  const { testGlobals, handlers, openControllerSelectCount } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.updateSessionBindings([
      {
        sourcePath: 'shortcuts.openControllerSelect',
        originalKey: 'Alt+D',
        key: { code: 'KeyD', modifiers: ['alt'] },
        actionType: 'session-action',
        actionId: 'openControllerSelect',
      },
    ] as never);

    testGlobals.dispatchKeydown({
      key: 'd',
      code: 'KeyD',
      altKey: true,
    });

    assert.equal(openControllerSelectCount(), 1);
    assert.deepEqual(testGlobals.sessionActions, []);
    assert.deepEqual(testGlobals.openedModalNotifications, ['controller-select']);
  } finally {
    testGlobals.restore();
  }
});

test('keyboard mode: configured controller debug binding opens locally without dispatching a session action', async () => {
  const { testGlobals, handlers, openControllerDebugCount } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.updateSessionBindings([
      {
        sourcePath: 'shortcuts.openControllerDebug',
        originalKey: 'Alt+Shift+D',
        key: { code: 'KeyD', modifiers: ['alt', 'shift'] },
        actionType: 'session-action',
        actionId: 'openControllerDebug',
      },
    ] as never);

    testGlobals.dispatchKeydown({
      key: 'D',
      code: 'KeyD',
      altKey: true,
      shiftKey: true,
    });

    assert.equal(openControllerDebugCount(), 1);
    assert.deepEqual(testGlobals.sessionActions, []);
    assert.deepEqual(testGlobals.openedModalNotifications, ['controller-debug']);
  } finally {
    testGlobals.restore();
  }
});

test('keyboard mode: configured controller debug binding is not swallowed while popup is visible', async () => {
  const { ctx, testGlobals, handlers, openControllerDebugCount } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    ctx.state.yomitanPopupVisible = true;
    testGlobals.setPopupVisible(true);
    handlers.updateSessionBindings([
      {
        sourcePath: 'shortcuts.openControllerDebug',
        originalKey: 'Alt+Shift+D',
        key: { code: 'KeyD', modifiers: ['alt', 'shift'] },
        actionType: 'session-action',
        actionId: 'openControllerDebug',
      },
    ] as never);

    testGlobals.dispatchKeydown({
      key: 'D',
      code: 'KeyD',
      altKey: true,
      shiftKey: true,
    });

    assert.equal(openControllerDebugCount(), 1);
    assert.deepEqual(testGlobals.sessionActions, []);
    assert.deepEqual(testGlobals.openedModalNotifications, ['controller-debug']);
  } finally {
    testGlobals.restore();
  }
});

test('keyboard mode: former fixed Alt+Shift+C does nothing when controller debug is remapped', async () => {
  const { testGlobals, handlers } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.updateSessionBindings([
      {
        sourcePath: 'shortcuts.openControllerDebug',
        originalKey: 'Alt+Shift+D',
        key: { code: 'KeyD', modifiers: ['alt', 'shift'] },
        actionType: 'session-action',
        actionId: 'openControllerDebug',
      },
    ] as never);

    testGlobals.dispatchKeydown({
      key: 'C',
      code: 'KeyC',
      altKey: true,
      shiftKey: true,
    });

    assert.deepEqual(testGlobals.sessionActions, []);
  } finally {
    testGlobals.restore();
  }
});

test('keyboard mode: controller select modal handles arrow keys before yomitan popup', async () => {
  const { ctx, testGlobals, handlers, controllerSelectKeydownCount } =
    createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    ctx.state.controllerSelectModalOpen = true;
    ctx.state.yomitanPopupVisible = true;
    testGlobals.setPopupVisible(true);

    testGlobals.dispatchKeydown({ key: 'ArrowDown', code: 'ArrowDown' });

    assert.equal(controllerSelectKeydownCount(), 1);
    assert.equal(
      testGlobals.commandEvents.some(
        (event) => event.type === 'forwardKeyDown' && event.code === 'ArrowDown',
      ),
      false,
    );
  } finally {
    testGlobals.restore();
  }
});

test('keyboard mode: playlist browser modal handles arrow keys before yomitan popup', async () => {
  const { ctx, testGlobals, handlers, playlistBrowserKeydownCount } =
    createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    ctx.state.playlistBrowserModalOpen = true;
    ctx.state.yomitanPopupVisible = true;
    testGlobals.setPopupVisible(true);

    testGlobals.dispatchKeydown({ key: 'ArrowDown', code: 'ArrowDown' });

    assert.equal(playlistBrowserKeydownCount(), 1);
    assert.equal(
      testGlobals.commandEvents.some(
        (event) => event.type === 'forwardKeyDown' && event.code === 'ArrowDown',
      ),
      false,
    );
  } finally {
    testGlobals.restore();
  }
});

test('keyboard mode: playlist browser modal handles h before lookup controls', async () => {
  const { ctx, testGlobals, handlers, playlistBrowserKeydownCount } =
    createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();
    ctx.state.playlistBrowserModalOpen = true;
    ctx.state.keyboardSelectedWordIndex = 2;

    testGlobals.dispatchKeydown({ key: 'h', code: 'KeyH' });

    assert.equal(playlistBrowserKeydownCount(), 1);
    assert.equal(ctx.state.keyboardSelectedWordIndex, 2);
  } finally {
    testGlobals.restore();
  }
});

test('keyboard mode: configured stats toggle works even while popup is open', async () => {
  const { handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    testGlobals.setPopupVisible(true);
    testGlobals.setStatsToggleKey('KeyG');
    await handlers.setupMpvInputForwarding();

    testGlobals.dispatchKeydown({ key: 'g', code: 'KeyG' });

    assert.equal(testGlobals.statsToggleOverlayCalls(), 1);
  } finally {
    testGlobals.restore();
  }
});

test('refreshConfiguredShortcuts updates refreshed stats and mark-watched keys', async () => {
  const { handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();

    testGlobals.setStatsToggleKey('KeyG');
    testGlobals.setMarkWatchedKey('KeyM');
    await handlers.refreshConfiguredShortcuts();

    const beforeMarkWatchedCalls = testGlobals.markActiveVideoWatchedCalls();

    testGlobals.dispatchKeydown({ key: 'g', code: 'KeyG' });
    testGlobals.dispatchKeydown({ key: 'm', code: 'KeyM' });
    await wait(10);

    assert.equal(testGlobals.statsToggleOverlayCalls(), 1);
    assert.equal(testGlobals.markActiveVideoWatchedCalls(), beforeMarkWatchedCalls + 1);
  } finally {
    testGlobals.restore();
  }
});

test('youtube picker: unhandled keys still dispatch mpv keybindings', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.updateSessionBindings([
      {
        sourcePath: 'keybindings[0].key',
        originalKey: 'Space',
        key: { code: 'Space', modifiers: [] },
        actionType: 'mpv-command',
        command: ['cycle', 'pause'],
      },
      {
        sourcePath: 'keybindings[1].key',
        originalKey: 'KeyQ',
        key: { code: 'KeyQ', modifiers: [] },
        actionType: 'mpv-command',
        command: ['quit'],
      },
    ] as never);

    ctx.state.youtubePickerModalOpen = true;

    testGlobals.dispatchKeydown({ key: ' ', code: 'Space' });
    testGlobals.dispatchKeydown({ key: 'q', code: 'KeyQ' });

    assert.deepEqual(testGlobals.mpvCommands.slice(-2), [['cycle', 'pause'], ['quit']]);
  } finally {
    testGlobals.restore();
  }
});

test('session binding: Ctrl+Alt+S dispatches subsync action locally', async () => {
  const { handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.updateSessionBindings([
      {
        sourcePath: 'shortcuts.triggerSubsync',
        originalKey: 'Ctrl+Alt+S',
        key: { code: 'KeyS', modifiers: ['ctrl', 'alt'] },
        actionType: 'session-action',
        actionId: 'triggerSubsync',
      },
    ] as never);

    testGlobals.dispatchKeydown({ key: 's', code: 'KeyS', ctrlKey: true, altKey: true });

    assert.deepEqual(testGlobals.sessionActions, [
      { actionId: 'triggerSubsync', payload: undefined },
    ]);
  } finally {
    testGlobals.restore();
  }
});

test('session binding: Ctrl+Shift+J dispatches jimaku action locally', async () => {
  const { handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.updateSessionBindings([
      {
        sourcePath: 'shortcuts.openJimaku',
        originalKey: 'Ctrl+Shift+J',
        key: { code: 'KeyJ', modifiers: ['ctrl', 'shift'] },
        actionType: 'session-action',
        actionId: 'openJimaku',
      },
    ] as never);

    testGlobals.dispatchKeydown({ key: 'J', code: 'KeyJ', ctrlKey: true, shiftKey: true });

    assert.deepEqual(testGlobals.sessionActions, [{ actionId: 'openJimaku', payload: undefined }]);
  } finally {
    testGlobals.restore();
  }
});

test('session binding: Ctrl+Shift+O dispatches runtime options locally', async () => {
  const { handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.updateSessionBindings([
      {
        sourcePath: 'shortcuts.openRuntimeOptions',
        originalKey: 'CommandOrControl+Shift+O',
        key: { code: 'KeyO', modifiers: ['ctrl', 'shift'] },
        actionType: 'session-action',
        actionId: 'openRuntimeOptions',
      },
    ] as never);

    testGlobals.dispatchKeydown({ key: 'O', code: 'KeyO', ctrlKey: true, shiftKey: true });

    assert.deepEqual(testGlobals.sessionActions, [
      { actionId: 'openRuntimeOptions', payload: undefined },
    ]);
  } finally {
    testGlobals.restore();
  }
});

test('session binding: copy subtitle multiple captures follow-up digit locally', async () => {
  const { handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.updateSessionBindings([
      {
        sourcePath: 'shortcuts.copySubtitleMultiple',
        originalKey: 'Ctrl+M',
        key: { code: 'KeyM', modifiers: ['ctrl'] },
        actionType: 'session-action',
        actionId: 'copySubtitleMultiple',
      },
    ] as never);

    testGlobals.dispatchKeydown({ key: 'm', code: 'KeyM', ctrlKey: true });
    testGlobals.dispatchKeydown({ key: '3', code: 'Digit3' });

    assert.deepEqual(testGlobals.sessionActions, [
      { actionId: 'copySubtitleMultiple', payload: { count: 3 } },
    ]);
  } finally {
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
    const openEvents = testGlobals.commandEvents.filter(
      (event) => event.type === 'scanSelectedText',
    );
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

    testGlobals.dispatchKeydown({ key: 'y', code: 'KeyY', ctrlKey: true });
    await wait(0);

    assert.equal(testGlobals.focusMainWindowCalls() > 0, true);
    assert.equal(testGlobals.windowFocusCalls() > 0, true);
    assert.equal(testGlobals.overlayFocusCalls.length > 0, true);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: visible-layer Ctrl+Shift+Y should not be toggled by renderer keydown', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    ctx.platform.isModalLayer = false;

    testGlobals.dispatchKeydown({ key: 'Y', code: 'KeyY', ctrlKey: true, shiftKey: true });
    assert.equal(ctx.state.keyboardDrivenModeEnabled, false);

    handlers.handleKeyboardModeToggleRequested();
    assert.equal(ctx.state.keyboardDrivenModeEnabled, true);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: modal-layer Ctrl+Shift+Y still toggles via renderer keydown', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    ctx.platform.isModalLayer = true;

    testGlobals.dispatchKeydown({ key: 'Y', code: 'KeyY', ctrlKey: true, shiftKey: true });
    assert.equal(ctx.state.keyboardDrivenModeEnabled, true);

    testGlobals.dispatchKeydown({ key: 'Y', code: 'KeyY', ctrlKey: true, shiftKey: true });
    assert.equal(ctx.state.keyboardDrivenModeEnabled, false);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: turning mode off clears selected token highlight', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    ctx.state.keyboardSelectedWordIndex = 1;
    handlers.syncKeyboardTokenSelection();
    const wordNodes = ctx.dom.subtitleRoot.querySelectorAll();
    assert.equal(wordNodes[1]?.classList.contains('keyboard-selected'), true);

    handlers.handleKeyboardModeToggleRequested();

    assert.equal(ctx.state.keyboardDrivenModeEnabled, false);
    assert.equal(ctx.state.keyboardSelectedWordIndex, null);
    assert.equal(wordNodes[1]?.classList.contains('keyboard-selected'), false);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: popup hidden after mode off clears stale selected token highlight', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    ctx.state.keyboardSelectedWordIndex = 1;
    ctx.state.yomitanPopupVisible = true;
    testGlobals.setPopupVisible(true);
    handlers.syncKeyboardTokenSelection();

    const wordNodes = ctx.dom.subtitleRoot.querySelectorAll();
    assert.equal(wordNodes[1]?.classList.contains('keyboard-selected'), true);

    handlers.handleKeyboardModeToggleRequested();
    ctx.state.yomitanPopupVisible = false;
    testGlobals.setPopupVisible(false);
    testGlobals.dispatchWindowEvent(YOMITAN_POPUP_HIDDEN_EVENT);

    assert.equal(ctx.state.keyboardDrivenModeEnabled, false);
    assert.equal(wordNodes[1]?.classList.contains('keyboard-selected'), false);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: closing lookup keeps controller selection but clears native text selection', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    ctx.state.keyboardSelectedWordIndex = 1;
    handlers.syncKeyboardTokenSelection();

    const wordNodes = ctx.dom.subtitleRoot.querySelectorAll();
    assert.equal(wordNodes[1]?.classList.contains('keyboard-selected'), true);
    assert.equal(ctx.dom.subtitleRoot.classList.contains('has-selection'), false);

    handlers.handleLookupWindowToggleRequested();
    await wait(0);
    assert.equal(ctx.dom.subtitleRoot.classList.contains('has-selection'), true);
    assert.equal(testGlobals.selectionAddCount() > 0, true);

    ctx.state.yomitanPopupVisible = true;
    testGlobals.setPopupVisible(true);
    handlers.closeLookupWindow();
    ctx.state.yomitanPopupVisible = false;
    testGlobals.setPopupVisible(false);
    testGlobals.dispatchWindowEvent(YOMITAN_POPUP_HIDDEN_EVENT);
    await wait(0);

    assert.equal(ctx.state.keyboardDrivenModeEnabled, true);
    assert.equal(wordNodes[1]?.classList.contains('keyboard-selected'), true);
    assert.equal(ctx.dom.subtitleRoot.classList.contains('has-selection'), false);
    assert.equal(testGlobals.selectionClearCount() > 0, true);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: closing lookup clears yomitan active text source so same token can reopen immediately', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    ctx.state.keyboardSelectedWordIndex = 1;
    handlers.syncKeyboardTokenSelection();

    handlers.handleLookupWindowToggleRequested();
    await wait(0);

    ctx.state.yomitanPopupVisible = true;
    testGlobals.setPopupVisible(true);
    handlers.handleLookupWindowToggleRequested();
    await wait(0);

    const closeCommands = testGlobals.commandEvents.filter(
      (event) => event.type === 'setVisible' || event.type === 'clearActiveTextSource',
    );
    assert.deepEqual(closeCommands.slice(-2), [
      { type: 'setVisible', visible: false },
      { type: 'clearActiveTextSource' },
    ]);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('subtitle refresh outside keyboard mode clears yomitan active text source', async () => {
  const { handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    handlers.handleSubtitleContentUpdated();
    await wait(0);

    const clearCommands = testGlobals.commandEvents.filter(
      (event) => event.type === 'clearActiveTextSource',
    );
    assert.deepEqual(clearCommands, [{ type: 'clearActiveTextSource' }]);
  } finally {
    testGlobals.restore();
  }
});

test('keyboard mode: lookup toggle closes popup when DOM visibility is the source of truth', async () => {
  const { ctx, handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    ctx.state.keyboardSelectedWordIndex = 1;
    handlers.syncKeyboardTokenSelection();
    ctx.state.yomitanPopupVisible = false;
    testGlobals.setPopupVisible(true);

    handlers.handleLookupWindowToggleRequested();
    await wait(0);

    const closeCommands = testGlobals.commandEvents.filter(
      (event) => event.type === 'setVisible' || event.type === 'clearActiveTextSource',
    );
    assert.deepEqual(closeCommands.slice(-2), [
      { type: 'setVisible', visible: false },
      { type: 'clearActiveTextSource' },
    ]);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: moving right beyond end jumps next subtitle and resets selector to start', async () => {
  const { ctx, handlers, testGlobals, setWordCount } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    setWordCount(3);
    ctx.state.keyboardSelectedWordIndex = 2;
    handlers.syncKeyboardTokenSelection();

    testGlobals.dispatchKeydown({ key: 'ArrowRight', code: 'ArrowRight' });
    await wait(0);
    assert.deepEqual(testGlobals.mpvCommands.at(-1), ['sub-seek', 1]);

    setWordCount(2);
    handlers.syncKeyboardTokenSelection();
    assert.equal(ctx.state.keyboardSelectedWordIndex, 0);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: moving left beyond start jumps previous subtitle and sets selector to end', async () => {
  const { ctx, handlers, testGlobals, setWordCount } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    setWordCount(3);
    ctx.state.keyboardSelectedWordIndex = 0;
    handlers.syncKeyboardTokenSelection();

    testGlobals.dispatchKeydown({ key: 'ArrowLeft', code: 'ArrowLeft' });
    await wait(0);
    assert.deepEqual(testGlobals.mpvCommands.at(-1), ['sub-seek', -1]);

    setWordCount(4);
    handlers.syncKeyboardTokenSelection();
    assert.equal(ctx.state.keyboardSelectedWordIndex, 3);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: empty subtitle gap left and right still seek adjacent subtitle lines', async () => {
  const { ctx, handlers, testGlobals, setWordCount } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    setWordCount(0);
    handlers.syncKeyboardTokenSelection();

    testGlobals.dispatchKeydown({ key: 'ArrowRight', code: 'ArrowRight' });
    await wait(0);
    assert.deepEqual(testGlobals.mpvCommands.at(-1), ['sub-seek', 1]);

    testGlobals.dispatchKeydown({ key: 'ArrowLeft', code: 'ArrowLeft' });
    await wait(0);
    assert.deepEqual(testGlobals.mpvCommands.at(-1), ['sub-seek', -1]);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('controller mode: empty subtitle gap horizontal move still seeks adjacent subtitle lines', async () => {
  const { ctx, handlers, testGlobals, setWordCount } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    setWordCount(0);
    handlers.syncKeyboardTokenSelection();

    assert.equal(handlers.moveSelectionForController(1), true);
    await wait(0);
    assert.deepEqual(testGlobals.mpvCommands.at(-1), ['sub-seek', 1]);

    assert.equal(handlers.moveSelectionForController(-1), true);
    await wait(0);
    assert.deepEqual(testGlobals.mpvCommands.at(-1), ['sub-seek', -1]);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: popup-open edge jump refreshes lookup on the new subtitle selection', async () => {
  const { ctx, handlers, testGlobals, setWordCount } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    setWordCount(2);
    ctx.state.keyboardSelectedWordIndex = 1;
    ctx.state.yomitanPopupVisible = true;
    testGlobals.setPopupVisible(true);
    handlers.syncKeyboardTokenSelection();

    testGlobals.dispatchKeydown({ key: 'ArrowRight', code: 'ArrowRight' });
    await wait(0);
    assert.deepEqual(testGlobals.mpvCommands.at(-1), ['sub-seek', 1]);

    setWordCount(3);
    handlers.syncKeyboardTokenSelection();
    await wait(80);

    assert.equal(ctx.state.keyboardSelectedWordIndex, 0);
    const openEvents = testGlobals.commandEvents.filter(
      (event) => event.type === 'scanSelectedText',
    );
    assert.equal(openEvents.length > 0, true);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: natural subtitle advance resets selector to the start of the new line', async () => {
  const { ctx, handlers, testGlobals, setWordCount } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    setWordCount(3);
    ctx.state.keyboardSelectedWordIndex = 2;
    handlers.syncKeyboardTokenSelection();

    handlers.handleSubtitleContentUpdated();
    setWordCount(4);
    handlers.syncKeyboardTokenSelection();

    assert.equal(ctx.state.keyboardSelectedWordIndex, 0);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: edge jump while paused re-applies paused state after subtitle seek', async () => {
  const { ctx, handlers, testGlobals, setWordCount } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    setWordCount(2);
    ctx.state.keyboardSelectedWordIndex = 1;
    handlers.syncKeyboardTokenSelection();
    testGlobals.setPlaybackPausedResponse(true);

    testGlobals.dispatchKeydown({ key: 'ArrowRight', code: 'ArrowRight' });
    await wait(0);

    assert.deepEqual(testGlobals.mpvCommands.slice(-2), [
      ['sub-seek', 1],
      ['set_property', 'pause', 'yes'],
    ]);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: left edge jump while paused re-applies paused state after subtitle seek', async () => {
  const { ctx, handlers, testGlobals, setWordCount } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    setWordCount(2);
    ctx.state.keyboardSelectedWordIndex = 0;
    handlers.syncKeyboardTokenSelection();
    testGlobals.setPlaybackPausedResponse(true);

    testGlobals.dispatchKeydown({ key: 'ArrowLeft', code: 'ArrowLeft' });
    await wait(0);

    assert.deepEqual(testGlobals.mpvCommands.slice(-2), [
      ['sub-seek', -1],
      ['set_property', 'pause', 'yes'],
    ]);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: h edge jump while paused re-applies paused state after subtitle seek', async () => {
  const { ctx, handlers, testGlobals, setWordCount } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    setWordCount(2);
    ctx.state.keyboardSelectedWordIndex = 0;
    handlers.syncKeyboardTokenSelection();
    testGlobals.setPlaybackPausedResponse(true);

    testGlobals.dispatchKeydown({ key: 'h', code: 'KeyH' });
    await wait(0);

    assert.deepEqual(testGlobals.mpvCommands.slice(-2), [
      ['sub-seek', -1],
      ['set_property', 'pause', 'yes'],
    ]);
  } finally {
    ctx.state.keyboardDrivenModeEnabled = false;
    testGlobals.restore();
  }
});

test('keyboard mode: edge jump with unknown pause state re-applies pause conservatively', async () => {
  const { ctx, handlers, testGlobals, setWordCount } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    handlers.handleKeyboardModeToggleRequested();

    setWordCount(2);
    ctx.state.keyboardSelectedWordIndex = 1;
    handlers.syncKeyboardTokenSelection();
    testGlobals.setPlaybackPausedResponse(null);

    testGlobals.dispatchKeydown({ key: 'ArrowRight', code: 'ArrowRight' });
    await wait(0);

    assert.deepEqual(testGlobals.mpvCommands.slice(-2), [
      ['sub-seek', 1],
      ['set_property', 'pause', 'yes'],
    ]);
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

test('mark-watched keybinding calls markActiveVideoWatched and sends mpv commands', async () => {
  const { handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    const beforeCalls = testGlobals.markActiveVideoWatchedCalls();
    const beforeMpvCount = testGlobals.mpvCommands.length;

    testGlobals.dispatchKeydown({ key: 'w', code: 'KeyW' });
    await wait(10);

    assert.equal(testGlobals.markActiveVideoWatchedCalls(), beforeCalls + 1);
    const newMpvCommands = testGlobals.mpvCommands.slice(beforeMpvCount);
    assert.deepEqual(newMpvCommands, [
      ['show-text', 'Marked as watched', '1500'],
      ['playlist-next', 'force'],
    ]);
  } finally {
    testGlobals.restore();
  }
});

test('mark-watched keybinding does not send mpv commands when no active session', async () => {
  const { handlers, testGlobals } = createKeyboardHandlerHarness();

  try {
    await handlers.setupMpvInputForwarding();
    testGlobals.setMarkActiveVideoWatchedResult(false);
    const beforeMpvCount = testGlobals.mpvCommands.length;

    testGlobals.dispatchKeydown({ key: 'w', code: 'KeyW' });
    await wait(10);

    assert.equal(testGlobals.markActiveVideoWatchedCalls() > 0, true);
    const newMpvCommands = testGlobals.mpvCommands.slice(beforeMpvCount);
    assert.deepEqual(newMpvCommands, []);
  } finally {
    testGlobals.restore();
  }
});

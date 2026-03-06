import assert from 'node:assert/strict';
import test from 'node:test';

import { createMouseHandlers } from './mouse.js';
import { YOMITAN_POPUP_HIDDEN_EVENT, YOMITAN_POPUP_SHOWN_EVENT } from '../yomitan-popup.js';

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

function createDeferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((nextResolve) => {
    resolve = nextResolve;
  });
  return { promise, resolve };
}

function waitForNextTick(): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, 0);
  });
}

function createMouseTestContext() {
  const overlayClassList = createClassList();
  const subtitleRootClassList = createClassList();
  const subtitleContainerClassList = createClassList();

  const ctx = {
    dom: {
      overlay: {
        classList: overlayClassList,
      },
      subtitleRoot: {
        classList: subtitleRootClassList,
      },
      subtitleContainer: {
        classList: subtitleContainerClassList,
        style: { cursor: '' },
        addEventListener: () => {},
      },
      secondarySubContainer: {
        addEventListener: () => {},
      },
    },
    platform: {
      shouldToggleMouseIgnore: false,
      isMacOSPlatform: false,
    },
    state: {
      isOverSubtitle: false,
      isDragging: false,
      dragStartY: 0,
      startYPercent: 0,
    },
  };

  return ctx;
}

test('auto-pause on subtitle hover pauses on enter and resumes on leave when enabled', async () => {
  const ctx = createMouseTestContext();
  const mpvCommands: Array<(string | number)[]> = [];

  const handlers = createMouseHandlers(ctx as never, {
    modalStateReader: {
      isAnySettingsModalOpen: () => false,
      isAnyModalOpen: () => false,
    },
    applyYPercent: () => {},
    getCurrentYPercent: () => 10,
    persistSubtitlePositionPatch: () => {},
    getSubtitleHoverAutoPauseEnabled: () => true,
    getYomitanPopupAutoPauseEnabled: () => false,
    getPlaybackPaused: async () => false,
    sendMpvCommand: (command) => {
      mpvCommands.push(command);
    },
  });

  await handlers.handleMouseEnter();
  await handlers.handleMouseLeave();

  assert.deepEqual(mpvCommands, [
    ['set_property', 'pause', 'yes'],
    ['set_property', 'pause', 'no'],
  ]);
});

test('auto-pause on subtitle hover skips when playback is already paused', async () => {
  const ctx = createMouseTestContext();
  const mpvCommands: Array<(string | number)[]> = [];

  const handlers = createMouseHandlers(ctx as never, {
    modalStateReader: {
      isAnySettingsModalOpen: () => false,
      isAnyModalOpen: () => false,
    },
    applyYPercent: () => {},
    getCurrentYPercent: () => 10,
    persistSubtitlePositionPatch: () => {},
    getSubtitleHoverAutoPauseEnabled: () => true,
    getYomitanPopupAutoPauseEnabled: () => false,
    getPlaybackPaused: async () => true,
    sendMpvCommand: (command) => {
      mpvCommands.push(command);
    },
  });

  await handlers.handleMouseEnter();
  await handlers.handleMouseLeave();

  assert.deepEqual(mpvCommands, []);
});

test('auto-pause on subtitle hover is skipped when disabled in config', async () => {
  const ctx = createMouseTestContext();
  const mpvCommands: Array<(string | number)[]> = [];

  const handlers = createMouseHandlers(ctx as never, {
    modalStateReader: {
      isAnySettingsModalOpen: () => false,
      isAnyModalOpen: () => false,
    },
    applyYPercent: () => {},
    getCurrentYPercent: () => 10,
    persistSubtitlePositionPatch: () => {},
    getSubtitleHoverAutoPauseEnabled: () => false,
    getYomitanPopupAutoPauseEnabled: () => false,
    getPlaybackPaused: async () => false,
    sendMpvCommand: (command) => {
      mpvCommands.push(command);
    },
  });

  await handlers.handleMouseEnter();
  await handlers.handleMouseLeave();

  assert.deepEqual(mpvCommands, []);
});

test('pending hover pause check is ignored when mouse leaves before pause state resolves', async () => {
  const ctx = createMouseTestContext();
  const mpvCommands: Array<(string | number)[]> = [];
  const deferred = createDeferred<boolean | null>();

  const handlers = createMouseHandlers(ctx as never, {
    modalStateReader: {
      isAnySettingsModalOpen: () => false,
      isAnyModalOpen: () => false,
    },
    applyYPercent: () => {},
    getCurrentYPercent: () => 10,
    persistSubtitlePositionPatch: () => {},
    getSubtitleHoverAutoPauseEnabled: () => true,
    getYomitanPopupAutoPauseEnabled: () => false,
    getPlaybackPaused: async () => deferred.promise,
    sendMpvCommand: (command) => {
      mpvCommands.push(command);
    },
  });

  const enterPromise = handlers.handleMouseEnter();
  await handlers.handleMouseLeave();
  deferred.resolve(false);
  await enterPromise;

  assert.deepEqual(mpvCommands, []);
});

test('hover pause resumes immediately on subtitle leave even when yomitan popup is visible', async () => {
  const ctx = createMouseTestContext();
  const mpvCommands: Array<(string | number)[]> = [];
  const previousWindow = (globalThis as { window?: unknown }).window;
  const previousDocument = (globalThis as { document?: unknown }).document;
  const previousMutationObserver = (globalThis as { MutationObserver?: unknown }).MutationObserver;
  const previousNode = (globalThis as { Node?: unknown }).Node;
  const windowListeners = new Map<string, Array<() => void>>();

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      addEventListener: (type: string, listener: () => void) => {
        const bucket = windowListeners.get(type) ?? [];
        bucket.push(listener);
        windowListeners.set(type, bucket);
      },
      electronAPI: {
        setIgnoreMouseEvents: () => {},
      },
      focus: () => {},
      innerHeight: 1000,
      getSelection: () => null,
      setTimeout,
      clearTimeout,
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      querySelector: () => null,
      querySelectorAll: () => [],
      body: {},
      elementFromPoint: () => null,
      addEventListener: () => {},
    },
  });
  Object.defineProperty(globalThis, 'MutationObserver', {
    configurable: true,
    value: class {
      observe() {}
    },
  });
  Object.defineProperty(globalThis, 'Node', {
    configurable: true,
    value: {
      ELEMENT_NODE: 1,
    },
  });

  try {
    const handlers = createMouseHandlers(ctx as never, {
      modalStateReader: {
        isAnySettingsModalOpen: () => false,
        isAnyModalOpen: () => false,
      },
      applyYPercent: () => {},
      getCurrentYPercent: () => 10,
      persistSubtitlePositionPatch: () => {},
      getSubtitleHoverAutoPauseEnabled: () => true,
      getYomitanPopupAutoPauseEnabled: () => false,
      getPlaybackPaused: async () => false,
      sendMpvCommand: (command) => {
        mpvCommands.push(command);
      },
    });

    handlers.setupYomitanObserver();
    for (const listener of windowListeners.get(YOMITAN_POPUP_SHOWN_EVENT) ?? []) {
      listener();
    }
    await handlers.handleMouseEnter();
    await handlers.handleMouseLeave();

    assert.deepEqual(mpvCommands, [
      ['set_property', 'pause', 'yes'],
      ['set_property', 'pause', 'no'],
    ]);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
    Object.defineProperty(globalThis, 'MutationObserver', {
      configurable: true,
      value: previousMutationObserver,
    });
    Object.defineProperty(globalThis, 'Node', { configurable: true, value: previousNode });
  }
});

test('auto-pause still works when yomitan popup is already visible', async () => {
  const ctx = createMouseTestContext();
  const mpvCommands: Array<(string | number)[]> = [];
  const previousWindow = (globalThis as { window?: unknown }).window;
  const previousDocument = (globalThis as { document?: unknown }).document;
  const previousMutationObserver = (globalThis as { MutationObserver?: unknown }).MutationObserver;
  const previousNode = (globalThis as { Node?: unknown }).Node;
  const windowListeners = new Map<string, Array<() => void>>();

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      addEventListener: (type: string, listener: () => void) => {
        const bucket = windowListeners.get(type) ?? [];
        bucket.push(listener);
        windowListeners.set(type, bucket);
      },
      electronAPI: {
        setIgnoreMouseEvents: () => {},
      },
      focus: () => {},
      innerHeight: 1000,
      getSelection: () => null,
      setTimeout,
      clearTimeout,
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      querySelector: () => null,
      querySelectorAll: () => [],
      body: {},
      elementFromPoint: () => null,
      addEventListener: () => {},
    },
  });
  Object.defineProperty(globalThis, 'MutationObserver', {
    configurable: true,
    value: class {
      observe() {}
    },
  });
  Object.defineProperty(globalThis, 'Node', {
    configurable: true,
    value: {
      ELEMENT_NODE: 1,
    },
  });

  try {
    const handlers = createMouseHandlers(ctx as never, {
      modalStateReader: {
        isAnySettingsModalOpen: () => false,
        isAnyModalOpen: () => false,
      },
      applyYPercent: () => {},
      getCurrentYPercent: () => 10,
      persistSubtitlePositionPatch: () => {},
      getSubtitleHoverAutoPauseEnabled: () => true,
      getYomitanPopupAutoPauseEnabled: () => false,
      getPlaybackPaused: async () => false,
      sendMpvCommand: (command) => {
        mpvCommands.push(command);
      },
    });

    handlers.setupYomitanObserver();
    for (const listener of windowListeners.get(YOMITAN_POPUP_SHOWN_EVENT) ?? []) {
      listener();
    }
    await handlers.handleMouseEnter();
    await handlers.handleMouseLeave();

    assert.deepEqual(mpvCommands, [
      ['set_property', 'pause', 'yes'],
      ['set_property', 'pause', 'no'],
    ]);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
    Object.defineProperty(globalThis, 'MutationObserver', {
      configurable: true,
      value: previousMutationObserver,
    });
    Object.defineProperty(globalThis, 'Node', { configurable: true, value: previousNode });
  }
});

test('popup open pauses and popup close resumes when yomitan popup auto-pause is enabled', async () => {
  const ctx = createMouseTestContext();
  const mpvCommands: Array<(string | number)[]> = [];
  const previousWindow = (globalThis as { window?: unknown }).window;
  const previousDocument = (globalThis as { document?: unknown }).document;
  const previousMutationObserver = (globalThis as { MutationObserver?: unknown }).MutationObserver;
  const previousNode = (globalThis as { Node?: unknown }).Node;
  const windowListeners = new Map<string, Array<() => void>>();

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      addEventListener: (type: string, listener: () => void) => {
        const bucket = windowListeners.get(type) ?? [];
        bucket.push(listener);
        windowListeners.set(type, bucket);
      },
      electronAPI: {
        setIgnoreMouseEvents: () => {},
      },
      focus: () => {},
      innerHeight: 1000,
      getSelection: () => null,
      setTimeout,
      clearTimeout,
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      querySelector: () => null,
      querySelectorAll: () => [],
      body: {},
      elementFromPoint: () => null,
      addEventListener: () => {},
    },
  });
  Object.defineProperty(globalThis, 'MutationObserver', {
    configurable: true,
    value: class {
      observe() {}
    },
  });
  Object.defineProperty(globalThis, 'Node', {
    configurable: true,
    value: {
      ELEMENT_NODE: 1,
    },
  });

  try {
    const handlers = createMouseHandlers(ctx as never, {
      modalStateReader: {
        isAnySettingsModalOpen: () => false,
        isAnyModalOpen: () => false,
      },
      applyYPercent: () => {},
      getCurrentYPercent: () => 10,
      persistSubtitlePositionPatch: () => {},
      getSubtitleHoverAutoPauseEnabled: () => true,
      getYomitanPopupAutoPauseEnabled: () => true,
      getPlaybackPaused: async () => false,
      sendMpvCommand: (command: (string | number)[]) => {
        mpvCommands.push(command);
      },
    });

    handlers.setupYomitanObserver();

    for (const listener of windowListeners.get(YOMITAN_POPUP_SHOWN_EVENT) ?? []) {
      listener();
    }
    await waitForNextTick();
    for (const listener of windowListeners.get(YOMITAN_POPUP_HIDDEN_EVENT) ?? []) {
      listener();
    }

    assert.deepEqual(mpvCommands, [
      ['set_property', 'pause', 'yes'],
      ['set_property', 'pause', 'no'],
    ]);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
    Object.defineProperty(globalThis, 'MutationObserver', {
      configurable: true,
      value: previousMutationObserver,
    });
    Object.defineProperty(globalThis, 'Node', { configurable: true, value: previousNode });
  }
});

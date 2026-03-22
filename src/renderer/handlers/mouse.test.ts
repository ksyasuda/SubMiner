import assert from 'node:assert/strict';
import test from 'node:test';

import type { SubtitleSidebarConfig } from '../../types';
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
  const secondarySubContainerClassList = createClassList();

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
        classList: secondarySubContainerClassList,
        addEventListener: () => {},
      },
    },
    platform: {
      shouldToggleMouseIgnore: false,
      isMacOSPlatform: false,
    },
    state: {
      isOverSubtitle: false,
      isOverSubtitleSidebar: false,
      subtitleSidebarModalOpen: false,
      subtitleSidebarConfig: null as SubtitleSidebarConfig | null,
      isDragging: false,
      dragStartY: 0,
      startYPercent: 0,
    },
  };

  return ctx;
}

test('secondary hover pauses on enter, reveals secondary subtitle, and resumes on leave when enabled', async () => {
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

  await handlers.handleSecondaryMouseEnter();
  assert.equal(ctx.dom.secondarySubContainer.classList.contains('secondary-sub-hover-active'), true);
  await handlers.handleSecondaryMouseLeave();
  assert.equal(ctx.dom.secondarySubContainer.classList.contains('secondary-sub-hover-active'), false);

  assert.deepEqual(mpvCommands, [
    ['set_property', 'pause', 'yes'],
    ['set_property', 'pause', 'no'],
  ]);
});

test('moving between primary and secondary subtitle containers keeps the hover pause active', async () => {
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

  await handlers.handleSecondaryMouseEnter();
  await handlers.handleSecondaryMouseLeave({
    relatedTarget: ctx.dom.subtitleContainer,
  } as unknown as MouseEvent);
  await handlers.handlePrimaryMouseEnter({
    relatedTarget: ctx.dom.secondarySubContainer,
  } as unknown as MouseEvent);

  assert.equal(ctx.state.isOverSubtitle, true);
  assert.deepEqual(mpvCommands, [['set_property', 'pause', 'yes']]);
});

test('secondary leave toward primary subtitle container clears the secondary hover class', async () => {
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

  await handlers.handleSecondaryMouseEnter();
  await handlers.handleSecondaryMouseLeave({
    relatedTarget: ctx.dom.subtitleContainer,
  } as unknown as MouseEvent);

  assert.equal(ctx.state.isOverSubtitle, false);
  assert.equal(ctx.dom.secondarySubContainer.classList.contains('secondary-sub-hover-active'), false);
  assert.deepEqual(mpvCommands, [['set_property', 'pause', 'yes']]);
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

test('primary hover pauses on enter without revealing secondary subtitle', async () => {
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

  await handlers.handlePrimaryMouseEnter();
  assert.equal(ctx.dom.secondarySubContainer.classList.contains('secondary-sub-hover-active'), false);
  await handlers.handlePrimaryMouseLeave();

  assert.deepEqual(mpvCommands, [
    ['set_property', 'pause', 'yes'],
    ['set_property', 'pause', 'no'],
  ]);
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

test('subtitle leave restores passthrough while embedded sidebar is open but not hovered', async () => {
  const ctx = createMouseTestContext();
  const ignoreMouseCalls: Array<[boolean, { forward?: boolean } | undefined]> = [];
  const previousWindow = (globalThis as { window?: unknown }).window;

  ctx.platform.shouldToggleMouseIgnore = true;
  ctx.state.isOverSubtitle = true;
  ctx.state.subtitleSidebarModalOpen = true;
  ctx.state.subtitleSidebarConfig = {
    enabled: true,
    autoOpen: false,
    layout: 'embedded',
    toggleKey: 'Backslash',
    pauseVideoOnHover: false,
    autoScroll: true,
    maxWidth: 360,
    opacity: 0.92,
    backgroundColor: 'rgba(54, 58, 79, 0.88)',
    textColor: '#cad3f5',
    fontFamily: '"Iosevka Aile", sans-serif',
    fontSize: 17,
    timestampColor: '#a5adcb',
    activeLineColor: '#f5bde6',
    activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
    hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
  };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
          ignoreMouseCalls.push([ignore, options]);
        },
      },
    },
  });

  try {
    const handlers = createMouseHandlers(ctx as never, {
      modalStateReader: {
        isAnySettingsModalOpen: () => false,
        isAnyModalOpen: () => true,
      },
      applyYPercent: () => {},
      getCurrentYPercent: () => 10,
      persistSubtitlePositionPatch: () => {},
      getSubtitleHoverAutoPauseEnabled: () => false,
      getYomitanPopupAutoPauseEnabled: () => false,
      getPlaybackPaused: async () => false,
      sendMpvCommand: () => {},
    });

    await handlers.handlePrimaryMouseLeave();

    assert.deepEqual(ignoreMouseCalls.at(-1), [true, { forward: true }]);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
  }
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

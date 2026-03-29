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
    toggle: (token: string, force?: boolean) => {
      if (force === undefined) {
        if (classes.has(token)) {
          classes.delete(token);
          return false;
        }
        classes.add(token);
        return true;
      }
      if (force) {
        classes.add(token);
        return true;
      }
      classes.delete(token);
      return false;
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
  assert.equal(
    ctx.dom.secondarySubContainer.classList.contains('secondary-sub-hover-active'),
    true,
  );
  await handlers.handleSecondaryMouseLeave();
  assert.equal(
    ctx.dom.secondarySubContainer.classList.contains('secondary-sub-hover-active'),
    false,
  );

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
  assert.equal(
    ctx.dom.secondarySubContainer.classList.contains('secondary-sub-hover-active'),
    false,
  );
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
  assert.equal(
    ctx.dom.secondarySubContainer.classList.contains('secondary-sub-hover-active'),
    false,
  );
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

test('restorePointerInteractionState reapplies the secondary hover class from pointer location', async () => {
  const ctx = createMouseTestContext();
  ctx.platform.shouldToggleMouseIgnore = true;

  const documentListeners = new Map<string, Array<(event: MouseEvent | PointerEvent) => void>>();
  const originalDocument = (globalThis as { document?: unknown }).document;
  const originalWindow = (globalThis as { window?: unknown }).window;

  const secondarySubContainer = ctx.dom.secondarySubContainer as unknown as object;
  const overlay = ctx.dom.overlay as unknown as { classList: ReturnType<typeof createClassList> };

  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      addEventListener: (type: string, listener: (event: MouseEvent | PointerEvent) => void) => {
        const listeners = documentListeners.get(type) ?? [];
        listeners.push(listener);
        documentListeners.set(type, listeners);
      },
      elementFromPoint: () => secondarySubContainer,
    },
  });
  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        setIgnoreMouseEvents: () => {},
      },
      innerHeight: 1000,
      getSelection: () => ({ rangeCount: 0, isCollapsed: true }),
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
      getSubtitleHoverAutoPauseEnabled: () => false,
      getYomitanPopupAutoPauseEnabled: () => false,
      getPlaybackPaused: async () => false,
      sendMpvCommand: () => {},
    });

    handlers.setupPointerTracking();
    await handlers.handleSecondaryMouseEnter({
      clientX: 10,
      clientY: 20,
    } as unknown as MouseEvent);
    handlers.restorePointerInteractionState();

    overlay.classList.add('interactive');
    const mousemove = documentListeners.get('mousemove')?.[0];
    assert.ok(mousemove);
    mousemove?.({ clientX: 10, clientY: 20 } as MouseEvent);

    assert.equal(ctx.state.isOverSubtitle, true);
    assert.equal(
      ctx.dom.secondarySubContainer.classList.contains('secondary-sub-hover-active'),
      true,
    );
  } finally {
    Object.defineProperty(globalThis, 'document', { configurable: true, value: originalDocument });
    Object.defineProperty(globalThis, 'window', { configurable: true, value: originalWindow });
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

test('restorePointerInteractionState re-enables subtitle hover when pointer is already over subtitles', () => {
  const ctx = createMouseTestContext();
  const originalWindow = globalThis.window;
  const originalDocument = globalThis.document;
  const ignoreCalls: Array<{ ignore: boolean; forward?: boolean }> = [];
  const documentListeners = new Map<string, Array<(event: unknown) => void>>();
  ctx.platform.shouldToggleMouseIgnore = true;

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
          ignoreCalls.push({ ignore, forward: options?.forward });
        },
      },
      getComputedStyle: () => ({
        visibility: 'hidden',
        display: 'none',
        opacity: '0',
      }),
      focus: () => {},
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      addEventListener: (type: string, listener: (event: unknown) => void) => {
        const bucket = documentListeners.get(type) ?? [];
        bucket.push(listener);
        documentListeners.set(type, bucket);
      },
      elementFromPoint: () => ctx.dom.subtitleContainer,
      querySelectorAll: () => [],
      body: {},
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
      getSubtitleHoverAutoPauseEnabled: () => false,
      getYomitanPopupAutoPauseEnabled: () => false,
      getPlaybackPaused: async () => false,
      sendMpvCommand: () => {},
    });

    handlers.setupPointerTracking();
    for (const listener of documentListeners.get('mousemove') ?? []) {
      listener({ clientX: 120, clientY: 240 });
    }
    handlers.restorePointerInteractionState();

    assert.equal(ctx.state.isOverSubtitle, true);
    assert.equal(ctx.dom.overlay.classList.contains('interactive'), true);
    assert.deepEqual(ignoreCalls, [
      { ignore: false, forward: undefined },
      { ignore: false, forward: undefined },
    ]);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: originalWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: originalDocument });
  }
});

test('pointer tracking enables overlay interaction as soon as the cursor reaches subtitles', () => {
  const ctx = createMouseTestContext();
  const originalWindow = globalThis.window;
  const originalDocument = globalThis.document;
  const ignoreCalls: Array<{ ignore: boolean; forward?: boolean }> = [];
  const documentListeners = new Map<string, Array<(event: unknown) => void>>();
  ctx.platform.shouldToggleMouseIgnore = true;

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
          ignoreCalls.push({ ignore, forward: options?.forward });
        },
      },
      getComputedStyle: () => ({
        visibility: 'hidden',
        display: 'none',
        opacity: '0',
      }),
      focus: () => {},
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      addEventListener: (type: string, listener: (event: unknown) => void) => {
        const bucket = documentListeners.get(type) ?? [];
        bucket.push(listener);
        documentListeners.set(type, bucket);
      },
      elementFromPoint: () => ctx.dom.subtitleContainer,
      querySelectorAll: () => [],
      body: {},
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
      getSubtitleHoverAutoPauseEnabled: () => false,
      getYomitanPopupAutoPauseEnabled: () => false,
      getPlaybackPaused: async () => false,
      sendMpvCommand: () => {},
    });

    handlers.setupPointerTracking();
    for (const listener of documentListeners.get('mousemove') ?? []) {
      listener({ clientX: 120, clientY: 240 });
    }

    assert.equal(ctx.state.isOverSubtitle, true);
    assert.equal(ctx.dom.overlay.classList.contains('interactive'), true);
    assert.deepEqual(ignoreCalls, [{ ignore: false, forward: undefined }]);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: originalWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: originalDocument });
  }
});

test('pointer tracking restores click-through after the cursor leaves subtitles', () => {
  const ctx = createMouseTestContext();
  const originalWindow = globalThis.window;
  const originalDocument = globalThis.document;
  const ignoreCalls: Array<{ ignore: boolean; forward?: boolean }> = [];
  const documentListeners = new Map<string, Array<(event: unknown) => void>>();
  let hoveredElement: unknown = ctx.dom.subtitleContainer;
  ctx.platform.shouldToggleMouseIgnore = true;

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
          ignoreCalls.push({ ignore, forward: options?.forward });
        },
      },
      getComputedStyle: () => ({
        visibility: 'hidden',
        display: 'none',
        opacity: '0',
      }),
      focus: () => {},
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      addEventListener: (type: string, listener: (event: unknown) => void) => {
        const bucket = documentListeners.get(type) ?? [];
        bucket.push(listener);
        documentListeners.set(type, bucket);
      },
      elementFromPoint: () => hoveredElement,
      querySelectorAll: () => [],
      body: {},
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
      getSubtitleHoverAutoPauseEnabled: () => false,
      getYomitanPopupAutoPauseEnabled: () => false,
      getPlaybackPaused: async () => false,
      sendMpvCommand: () => {},
    });

    handlers.setupPointerTracking();
    for (const listener of documentListeners.get('mousemove') ?? []) {
      listener({ clientX: 120, clientY: 240 });
    }

    hoveredElement = null;
    for (const listener of documentListeners.get('mousemove') ?? []) {
      listener({ clientX: 640, clientY: 360 });
    }

    assert.equal(ctx.state.isOverSubtitle, false);
    assert.equal(ctx.dom.overlay.classList.contains('interactive'), false);
    assert.deepEqual(ignoreCalls, [
      { ignore: false, forward: undefined },
      { ignore: true, forward: true },
    ]);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: originalWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: originalDocument });
  }
});

test('restorePointerInteractionState keeps overlay interactive until first real pointer move can resync hover', () => {
  const ctx = createMouseTestContext();
  const originalWindow = globalThis.window;
  const originalDocument = globalThis.document;
  const ignoreCalls: Array<{ ignore: boolean; forward?: boolean }> = [];
  const documentListeners = new Map<string, Array<(event: unknown) => void>>();
  let hoveredElement: unknown = null;
  ctx.platform.shouldToggleMouseIgnore = true;

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
          ignoreCalls.push({ ignore, forward: options?.forward });
        },
      },
      getComputedStyle: () => ({
        visibility: 'hidden',
        display: 'none',
        opacity: '0',
      }),
      focus: () => {},
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      addEventListener: (type: string, listener: (event: unknown) => void) => {
        const bucket = documentListeners.get(type) ?? [];
        bucket.push(listener);
        documentListeners.set(type, bucket);
      },
      elementFromPoint: () => hoveredElement,
      querySelectorAll: () => [],
      body: {},
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
      getSubtitleHoverAutoPauseEnabled: () => false,
      getYomitanPopupAutoPauseEnabled: () => false,
      getPlaybackPaused: async () => false,
      sendMpvCommand: () => {},
    });

    handlers.setupPointerTracking();
    handlers.restorePointerInteractionState();

    assert.equal(ctx.state.isOverSubtitle, false);
    assert.equal(ctx.dom.overlay.classList.contains('interactive'), true);
    assert.deepEqual(ignoreCalls, [
      { ignore: true, forward: true },
      { ignore: false, forward: undefined },
    ]);

    hoveredElement = null;
    for (const listener of documentListeners.get('mousemove') ?? []) {
      listener({ clientX: 24, clientY: 48 });
    }

    assert.equal(ctx.state.isOverSubtitle, false);
    assert.equal(ctx.dom.overlay.classList.contains('interactive'), false);
    assert.deepEqual(ignoreCalls, [
      { ignore: true, forward: true },
      { ignore: false, forward: undefined },
      { ignore: true, forward: true },
    ]);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: originalWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: originalDocument });
  }
});

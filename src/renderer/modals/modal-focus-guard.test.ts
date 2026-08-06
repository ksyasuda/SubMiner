import assert from 'node:assert/strict';
import test from 'node:test';

import { createModalFocusGuard } from './modal-focus-guard';

type Listener = (event?: unknown) => void;

function createRoot(contains: boolean) {
  const listeners: Array<{ type: string; listener: Listener }> = [];
  return {
    contains: () => contains,
    addEventListener: (type: string, listener: Listener) => {
      listeners.push({ type, listener });
    },
    removeEventListener: (type: string, listener: Listener) => {
      const index = listeners.findIndex(
        (entry) => entry.type === type && entry.listener === listener,
      );
      if (index >= 0) listeners.splice(index, 1);
    },
    listeners,
  };
}

type Harness = {
  guard: ReturnType<typeof createModalFocusGuard>;
  root: ReturnType<typeof createRoot>;
  focusMainWindowCalls: () => number;
  focused: () => string[];
  documentListeners: () => string[];
  windowListeners: () => string[];
  handlerFor: (scope: 'document' | 'window', type: string) => Listener | undefined;
  setActiveElement: (value: unknown) => void;
  advanceClock: (ms: number) => void;
  runTimers: () => void;
  clearedTimers: () => number[];
  pendingTimerCount: () => number;
  restore: () => void;
};

function createHarness(
  options: {
    isOpen?: () => boolean;
    isModalLayer?: boolean;
    contains?: boolean;
    preferredVisible?: boolean;
    preferredAcceptsFocus?: boolean;
    extraPreferred?: boolean;
    fallback?: 'element' | null;
  } = {},
): Harness {
  const previous = (['window', 'document', 'HTMLElement', 'Element'] as const).map(
    (name) => [name, Object.getOwnPropertyDescriptor(globalThis, name)] as const,
  );

  class TestElement {}
  for (const name of ['HTMLElement', 'Element'] as const) {
    Object.defineProperty(globalThis, name, {
      configurable: true,
      writable: true,
      value: TestElement,
    });
  }

  let focusMainWindowCalls = 0;
  let now = 1_000;
  const realDateNow = Date.now;
  Date.now = () => now;
  const focused: string[] = [];
  const documentListeners: Array<{ type: string; listener: Listener }> = [];
  const windowListeners: Array<{ type: string; listener: Listener }> = [];

  const register = (registry: Array<{ type: string; listener: Listener }>) => ({
    add: (type: string, listener: Listener) => {
      registry.push({ type, listener });
    },
    remove: (type: string, listener: Listener) => {
      const index = registry.findIndex(
        (entry) => entry.type === type && entry.listener === listener,
      );
      if (index >= 0) registry.splice(index, 1);
    },
  });
  const windowRegistry = register(windowListeners);
  const documentRegistry = register(documentListeners);
  const timers = new Map<number, () => void>();
  let nextTimerId = 0;
  const clearedTimers: number[] = [];
  let activeElement: unknown = null;

  const root = createRoot(options.contains ?? false);

  const preferred = Object.assign(new TestElement(), {
    // A position:fixed element has a null offsetParent but still has rects.
    offsetParent: null,
    getClientRects: () => (options.preferredVisible === false ? [] : [{ width: 10, height: 10 }]),
    focus: () => {
      focused.push('preferred');
      // A rendered-but-unfocusable target (e.g. disabled) never becomes active.
      if (options.preferredAcceptsFocus !== false) activeElement = preferred;
    },
  });
  const secondPreferred = Object.assign(new TestElement(), {
    getClientRects: () => [{ width: 10, height: 10 }],
    focus: () => {
      focused.push('second');
      activeElement = secondPreferred;
    },
  });
  const fallback =
    options.fallback === null
      ? null
      : Object.assign(new TestElement(), {
          focus: () => {
            focused.push('fallback');
            activeElement = fallback;
          },
        });

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    writable: true,
    value: {
      electronAPI: {
        focusMainWindow: async () => {
          focusMainWindowCalls += 1;
        },
      },
      focus: () => {
        focused.push('window');
      },
      addEventListener: windowRegistry.add,
      removeEventListener: windowRegistry.remove,
      setTimeout: (callback: () => void) => {
        nextTimerId += 1;
        timers.set(nextTimerId, callback);
        return nextTimerId;
      },
      clearTimeout: (id: number) => {
        clearedTimers.push(id);
        timers.delete(id);
      },
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    writable: true,
    value: {
      get activeElement() {
        return activeElement;
      },
      addEventListener: documentRegistry.add,
      removeEventListener: documentRegistry.remove,
    },
  });

  const guard = createModalFocusGuard({
    isOpen: options.isOpen ?? (() => true),
    getModalRoot: () => root as unknown as Element,
    getPreferredFocusTargets: () =>
      (options.extraPreferred
        ? [preferred, secondPreferred]
        : [preferred]) as unknown as HTMLElement[],
    getFallbackFocusTarget: () => fallback as unknown as Element | null,
    isModalLayer: options.isModalLayer ?? true,
  });

  return {
    guard,
    root,
    focusMainWindowCalls: () => focusMainWindowCalls,
    focused: () => focused,
    documentListeners: () => documentListeners.map((entry) => entry.type),
    windowListeners: () => windowListeners.map((entry) => entry.type),
    handlerFor: (scope: 'document' | 'window', type: string) =>
      (scope === 'document' ? documentListeners : windowListeners).find(
        (entry) => entry.type === type,
      )?.listener,
    setActiveElement: (value: unknown) => {
      activeElement = value;
    },
    advanceClock: (ms: number) => {
      now += ms;
    },
    runTimers: () => {
      const pending = [...timers.values()];
      timers.clear();
      for (const callback of pending) callback();
    },
    clearedTimers: () => clearedTimers,
    pendingTimerCount: () => timers.size,
    restore: () => {
      Date.now = realDateNow;
      for (const [name, descriptor] of previous) {
        if (descriptor) {
          Object.defineProperty(globalThis, name, descriptor);
        } else {
          delete (globalThis as Record<string, unknown>)[name];
        }
      }
    },
  };
}

test('modal focus guard attaches once and detaches every listener', () => {
  const harness = createHarness();
  try {
    harness.guard.attach();
    harness.guard.attach();

    assert.deepEqual(harness.documentListeners(), ['focusin']);
    assert.deepEqual(harness.windowListeners(), ['blur', 'focus']);
    assert.deepEqual(
      harness.root.listeners.map((entry) => entry.type),
      ['pointerdown', 'click'],
    );

    harness.guard.detach();

    assert.deepEqual(harness.documentListeners(), []);
    assert.deepEqual(harness.windowListeners(), []);
    assert.deepEqual(harness.root.listeners, []);
  } finally {
    harness.restore();
  }
});

test('modal focus guard pulls focus back when focusin lands outside the modal', () => {
  const harness = createHarness();
  try {
    harness.guard.attach();

    const focusin = harness.handlerFor('document', 'focusin');
    assert.ok(focusin, 'attach registers a focusin handler');

    // focusin is not cancelable, so recovery is the only observable effect.
    focusin?.({ target: {} });
    assert.deepEqual(harness.focused(), ['preferred']);
  } finally {
    harness.restore();
  }
});

test('modal focus guard ignores focusin while the modal is closed', () => {
  const harness = createHarness({ isOpen: () => false });
  try {
    harness.guard.attach();
    harness.handlerFor('document', 'focusin')?.({ target: {} });

    assert.deepEqual(harness.focused(), []);
  } finally {
    harness.restore();
  }
});

test('modal focus guard restores focus to the first rendered target', () => {
  // The target is position:fixed (null offsetParent) yet visible, so it must
  // still win over the fallback.
  const harness = createHarness();
  try {
    harness.guard.enforceModalFocus();
    assert.deepEqual(harness.focused(), ['preferred']);
  } finally {
    harness.restore();
  }
});

test('modal focus guard tries the next target when one refuses focus', () => {
  const harness = createHarness({ preferredAcceptsFocus: false, extraPreferred: true });
  try {
    assert.equal(harness.guard.focusFallbackTarget(), true);
    assert.deepEqual(harness.focused(), ['preferred', 'second']);
  } finally {
    harness.restore();
  }
});

test('modal focus guard reaches the fallback when no preferred target takes focus', () => {
  const harness = createHarness({ preferredAcceptsFocus: false });
  try {
    assert.equal(harness.guard.focusFallbackTarget(), true);
    assert.deepEqual(harness.focused(), ['preferred', 'fallback']);
  } finally {
    harness.restore();
  }
});

test('modal focus guard falls back when no preferred target is rendered', () => {
  const harness = createHarness({ preferredVisible: false });
  try {
    harness.guard.enforceModalFocus();
    assert.deepEqual(harness.focused(), ['fallback']);
  } finally {
    harness.restore();
  }
});

test('modal focus guard focuses the window when nothing else can take focus', () => {
  const harness = createHarness({ preferredVisible: false, fallback: null });
  try {
    assert.equal(harness.guard.focusFallbackTarget(), false);
    assert.deepEqual(harness.focused(), ['window']);
  } finally {
    harness.restore();
  }
});

test('modal focus guard leaves focus alone while it is already inside the modal', () => {
  const harness = createHarness({ contains: true });
  try {
    harness.setActiveElement(
      Object.create((globalThis as { Element: { prototype: object } }).Element.prototype),
    );
    harness.guard.enforceModalFocus();

    assert.deepEqual(harness.focused(), []);
  } finally {
    harness.restore();
  }
});

test('modal focus guard does nothing while the modal is closed', () => {
  const harness = createHarness({ isOpen: () => false });
  try {
    harness.guard.enforceModalFocus();
    assert.deepEqual(harness.focused(), []);
  } finally {
    harness.restore();
  }
});

test('modal focus guard clears recovery state on detach so a reopen is not blocked', () => {
  const harness = createHarness();
  try {
    harness.guard.attach();
    harness.guard.enforceModalFocus();
    assert.deepEqual(harness.focused(), ['preferred']);
    assert.equal(harness.pendingTimerCount(), 1, 'recovery armed the debounce timer');

    // Close while recovery is still in flight.
    harness.guard.detach();
    assert.equal(harness.clearedTimers().length, 1, 'the pending debounce timer is cancelled');
    assert.equal(harness.pendingTimerCount(), 0);

    // Immediate reopen: recovery must work straight away, not 120 ms later.
    harness.guard.attach();
    harness.setActiveElement(null);
    harness.guard.enforceModalFocus();

    assert.deepEqual(harness.focused(), ['preferred', 'preferred']);
  } finally {
    harness.restore();
  }
});

test('modal focus guard debounces recovery so a focus fight cannot spin', () => {
  const harness = createHarness();
  try {
    harness.guard.enforceModalFocus();
    harness.setActiveElement(null);

    // Re-entry guard: the recovery timer has not fired yet.
    harness.guard.enforceModalFocus();
    assert.deepEqual(harness.focused(), ['preferred']);

    // Timer cleared the re-entry flag, but the debounce window still holds.
    harness.runTimers();
    harness.guard.enforceModalFocus();
    assert.deepEqual(harness.focused(), ['preferred'], 'debounce still blocks the retry');

    // Past the window, recovery resumes.
    harness.advanceClock(200);
    harness.setActiveElement(null);
    harness.guard.enforceModalFocus();
    assert.deepEqual(harness.focused(), ['preferred', 'preferred']);
  } finally {
    harness.restore();
  }
});

test('modal focus guard asks the main window for focus off the modal layer only', () => {
  const onModalLayer = createHarness({ isModalLayer: true });
  try {
    onModalLayer.guard.requestOverlayFocus();
    assert.equal(onModalLayer.focusMainWindowCalls(), 0);
  } finally {
    onModalLayer.restore();
  }

  const onOverlayLayer = createHarness({ isModalLayer: false });
  try {
    onOverlayLayer.guard.requestOverlayFocus();
    assert.equal(onOverlayLayer.focusMainWindowCalls(), 1);
  } finally {
    onOverlayLayer.restore();
  }
});

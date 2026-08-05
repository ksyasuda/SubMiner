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
  setActiveElement: (value: unknown) => void;
  advanceClock: (ms: number) => void;
  runTimers: () => void;
  restore: () => void;
};

function createHarness(
  options: {
    isOpen?: () => boolean;
    isModalLayer?: boolean;
    contains?: boolean;
    preferredVisible?: boolean;
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
  const documentListeners: string[] = [];
  const windowListeners: string[] = [];
  const timers: Array<() => void> = [];
  let activeElement: unknown = null;

  const root = createRoot(options.contains ?? false);

  const preferred = Object.assign(new TestElement(), {
    // A position:fixed element has a null offsetParent but still has rects.
    offsetParent: null,
    getClientRects: () => (options.preferredVisible === false ? [] : [{ width: 10, height: 10 }]),
    focus: () => {
      focused.push('preferred');
      activeElement = preferred;
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
      addEventListener: (type: string) => {
        windowListeners.push(type);
      },
      removeEventListener: (type: string) => {
        const index = windowListeners.indexOf(type);
        if (index >= 0) windowListeners.splice(index, 1);
      },
      setTimeout: (callback: () => void) => {
        timers.push(callback);
        return timers.length;
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
      addEventListener: (type: string) => {
        documentListeners.push(type);
      },
      removeEventListener: (type: string) => {
        const index = documentListeners.indexOf(type);
        if (index >= 0) documentListeners.splice(index, 1);
      },
    },
  });

  const guard = createModalFocusGuard({
    isOpen: options.isOpen ?? (() => true),
    getModalRoot: () => root as unknown as Element,
    getPreferredFocusTargets: () => [preferred as unknown as HTMLElement],
    getFallbackFocusTarget: () => fallback as unknown as Element | null,
    isModalLayer: options.isModalLayer ?? true,
  });

  return {
    guard,
    root,
    focusMainWindowCalls: () => focusMainWindowCalls,
    focused: () => focused,
    documentListeners: () => documentListeners,
    windowListeners: () => windowListeners,
    setActiveElement: (value: unknown) => {
      activeElement = value;
    },
    advanceClock: (ms: number) => {
      now += ms;
    },
    runTimers: () => {
      while (timers.length > 0) timers.shift()?.();
    },
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

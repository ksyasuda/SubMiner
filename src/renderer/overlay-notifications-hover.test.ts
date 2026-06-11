import assert from 'node:assert/strict';
import test from 'node:test';

import { createOverlayNotificationRenderer } from './overlay-notifications';

function createClassList() {
  const tokens = new Set<string>();
  return {
    add: (...entries: string[]) => {
      for (const entry of entries) tokens.add(entry);
    },
    remove: (...entries: string[]) => {
      for (const entry of entries) tokens.delete(entry);
    },
    contains: (entry: string) => tokens.has(entry),
    toggle: (entry: string, force?: boolean) => {
      if (force === true) tokens.add(entry);
      else if (force === false) tokens.delete(entry);
      else if (tokens.has(entry)) tokens.delete(entry);
      else tokens.add(entry);
    },
  };
}

type FakeElement = {
  tagName: string;
  className: string;
  textContent: string;
  type: string;
  dataset: Record<string, string>;
  children: FakeElement[];
  classList: ReturnType<typeof createClassList>;
  append: (...children: FakeElement[]) => void;
  replaceChildren: (...children: FakeElement[]) => void;
  remove: () => void;
  setAttribute: (name: string, value: string) => void;
  addEventListener: (type: string, listener: (event?: unknown) => void) => void;
  dispatchEventType: (type: string, event?: unknown) => void;
};

function createFakeElement(tagName = 'div'): FakeElement {
  const listeners = new Map<string, Array<(event?: unknown) => void>>();
  const element: FakeElement = {
    tagName: tagName.toUpperCase(),
    className: '',
    textContent: '',
    type: '',
    dataset: {},
    children: [],
    classList: createClassList(),
    append: (...children) => {
      for (const child of children) {
        const existingIndex = element.children.indexOf(child);
        if (existingIndex >= 0) {
          element.children.splice(existingIndex, 1);
        }
        element.children.push(child);
      }
    },
    replaceChildren: (...children) => {
      element.children = [...children];
    },
    setAttribute: () => undefined,
    remove: () => undefined,
    addEventListener: (type, listener) => {
      listeners.set(type, [...(listeners.get(type) ?? []), listener]);
    },
    dispatchEventType: (type, event) => {
      for (const listener of listeners.get(type) ?? []) listener(event);
    },
  };
  return element;
}

function findChildByClass(element: FakeElement, className: string): FakeElement | null {
  if (element.className.split(/\s+/).includes(className)) {
    return element;
  }
  for (const child of element.children) {
    const match = findChildByClass(child, className);
    if (match) return match;
  }
  return null;
}

function createHoverContext(stack: FakeElement, ignoreCalls: Array<{ ignore: boolean }>) {
  return {
    dom: {
      overlay: { classList: createClassList() },
      overlayNotificationStack: stack,
    },
    platform: {
      shouldToggleMouseIgnore: true,
    },
    state: {
      isOverSubtitle: false,
      isOverSubtitleSidebar: false,
      isOverOverlayNotification: false,
      isOverNotificationHistory: false,
      yomitanPopupVisible: false,
      controllerSelectModalOpen: false,
      controllerDebugModalOpen: false,
      jimakuModalOpen: false,
      youtubePickerModalOpen: false,
      kikuModalOpen: false,
      runtimeOptionsModalOpen: false,
      subsyncModalOpen: false,
      sessionHelpModalOpen: false,
    },
  };
}

function installDomGlobals(ignoreCalls: Array<{ ignore: boolean }>): () => void {
  const originalDocument = Object.getOwnPropertyDescriptor(globalThis, 'document');
  const originalWindow = Object.getOwnPropertyDescriptor(globalThis, 'window');

  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    writable: true,
    value: {
      createElement: (tagName: string) => createFakeElement(tagName),
      querySelectorAll: () => [],
    },
  });
  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    writable: true,
    value: {
      clearTimeout: () => undefined,
      setTimeout: () => 1,
      electronAPI: {
        setIgnoreMouseEvents: (ignore: boolean) => {
          ignoreCalls.push({ ignore });
        },
      },
    },
  });

  return () => {
    if (originalDocument) {
      Object.defineProperty(globalThis, 'document', originalDocument);
    } else {
      delete (globalThis as { document?: unknown }).document;
    }
    if (originalWindow) {
      Object.defineProperty(globalThis, 'window', originalWindow);
    } else {
      delete (globalThis as { window?: unknown }).window;
    }
  };
}

test('passive overlay notification hover stays click-through on macOS passthrough overlays', () => {
  const stack = createFakeElement();
  const ignoreCalls: Array<{ ignore: boolean }> = [];
  const ctx = createHoverContext(stack, ignoreCalls);
  const restore = installDomGlobals(ignoreCalls);

  try {
    const renderer = createOverlayNotificationRenderer(ctx as never);
    renderer.show({
      id: 'character-dictionary-auto-sync',
      title: 'Character dictionary',
      body: 'Building character dictionary...',
      variant: 'progress',
      persistent: true,
    });

    stack.dispatchEventType('mouseenter');

    assert.equal(ctx.state.isOverOverlayNotification, false);
    assert.deepEqual(ignoreCalls, []);

    const card = stack.children[0];
    const close = card ? findChildByClass(card, 'overlay-notification-close') : null;
    if (!close) {
      assert.fail('Expected overlay notification close button.');
    }

    close.dispatchEventType('mouseenter');
    assert.equal(ctx.state.isOverOverlayNotification, false);
    assert.deepEqual(ignoreCalls, []);
  } finally {
    restore();
  }
});

test('overlay notification controls become interactive on hover', () => {
  const stack = createFakeElement();
  const ignoreCalls: Array<{ ignore: boolean }> = [];
  const ctx = createHoverContext(stack, ignoreCalls);
  const restore = installDomGlobals(ignoreCalls);

  try {
    const renderer = createOverlayNotificationRenderer(ctx as never);
    renderer.show({
      id: 'mined-card',
      title: 'Card created',
      body: 'Added sentence card',
      actions: [{ id: 'open-anki-card', label: 'Open in Anki', noteId: 42 }],
      persistent: true,
    });

    const card = stack.children[0];
    const action = card ? findChildByClass(card, 'overlay-notification-action') : null;
    if (!action) {
      assert.fail('Expected overlay notification action.');
    }

    action.dispatchEventType('mouseenter');
    assert.equal(ctx.state.isOverOverlayNotification, true);
    assert.deepEqual(ignoreCalls, [{ ignore: false }]);

    action.dispatchEventType('mouseleave');
    assert.equal(ctx.state.isOverOverlayNotification, false);
    assert.deepEqual(ignoreCalls, [{ ignore: false }, { ignore: true }]);
  } finally {
    restore();
  }
});

test('action overlay notification stack hover keeps card controls interactive', () => {
  const stack = createFakeElement();
  const ignoreCalls: Array<{ ignore: boolean }> = [];
  const ctx = createHoverContext(stack, ignoreCalls);
  const restore = installDomGlobals(ignoreCalls);

  try {
    const renderer = createOverlayNotificationRenderer(ctx as never);
    renderer.show({
      id: 'anki-card-updated',
      title: 'Anki Card Updated',
      body: 'Updated card: 食べる',
      persistent: true,
      actions: [{ id: 'open-anki-card', label: 'Open in Anki', noteId: 42 }],
    });

    stack.dispatchEventType('mouseenter');

    assert.equal(ctx.state.isOverOverlayNotification, true);
    assert.deepEqual(ignoreCalls, [{ ignore: false }]);
  } finally {
    restore();
  }
});

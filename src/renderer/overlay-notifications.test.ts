import assert from 'node:assert/strict';
import test from 'node:test';

import {
  createOverlayNotificationRenderer,
  createOverlayNotificationStore,
  handleOverlayNotificationEvent,
} from './overlay-notifications';

function createClassList(initialTokens: string[] = []) {
  const tokens = new Set(initialTokens);
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
  src: string;
  alt: string;
  type: string;
  dataset: Record<string, string>;
  children: FakeElement[];
  classList: ReturnType<typeof createClassList>;
  appendCalls: number;
  replaceChildrenCalls: number;
  append: (...children: FakeElement[]) => void;
  replaceChildren: (...children: FakeElement[]) => void;
  remove: () => void;
  setAttribute: (name: string, value: string) => void;
  getAttribute: (name: string) => string | null;
  addEventListener: (type: string, listener: (event?: unknown) => void) => void;
  dispatchEventType: (type: string, event?: unknown) => void;
};

function createFakeElement(tagName = 'div'): FakeElement {
  const attributes = new Map<string, string>();
  const listeners = new Map<string, Array<(event?: unknown) => void>>();
  const element: FakeElement = {
    tagName: tagName.toUpperCase(),
    className: '',
    textContent: '',
    src: '',
    alt: '',
    type: '',
    dataset: {},
    children: [],
    classList: createClassList(),
    appendCalls: 0,
    replaceChildrenCalls: 0,
    append: (...children) => {
      element.appendCalls += 1;
      for (const child of children) {
        const existingIndex = element.children.indexOf(child);
        if (existingIndex >= 0) {
          element.children.splice(existingIndex, 1);
        }
        element.children.push(child);
      }
    },
    replaceChildren: (...children) => {
      element.replaceChildrenCalls += 1;
      element.children = [...children];
    },
    setAttribute: (name, value) => {
      attributes.set(name, value);
    },
    getAttribute: (name) => attributes.get(name) ?? null,
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

type FakeWindow = {
  clearTimeout: (id: number) => void;
  setTimeout: (callback: () => void, delayMs: number) => number;
  electronAPI?: Record<string, unknown>;
};

// Installs a fake `document`/`window`, builds a renderer over a fake stack, runs `fn`, then
// restores the globals.
function withRenderer(
  windowOverrides: Partial<FakeWindow>,
  fn: (env: {
    renderer: ReturnType<typeof createOverlayNotificationRenderer>;
    stack: FakeElement;
  }) => void,
): void {
  const originalDocument = Object.getOwnPropertyDescriptor(globalThis, 'document');
  const originalWindow = Object.getOwnPropertyDescriptor(globalThis, 'window');
  const fakeWindow: FakeWindow = {
    clearTimeout: () => undefined,
    setTimeout: () => 1,
    ...windowOverrides,
  };
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    writable: true,
    value: { createElement: (tagName: string) => createFakeElement(tagName) },
  });
  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    writable: true,
    value: fakeWindow,
  });

  try {
    const stack = createFakeElement();
    const renderer = createOverlayNotificationRenderer({
      dom: { overlayNotificationStack: stack },
      state: { isOverOverlayNotification: false },
    } as never);
    fn({ renderer, stack });
  } finally {
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
  }
}

function requireChild(element: FakeElement | undefined, className?: string): FakeElement {
  const match = element && className ? findChildByClass(element, className) : element;
  if (!match) {
    assert.fail(`Expected element${className ? ` .${className}` : ''}.`);
  }
  return match;
}

test('overlay notification store caps transient notifications and keeps pinned jobs visible', () => {
  const store = createOverlayNotificationStore({ maxVisible: 3 });

  store.upsert({
    id: 'character-dictionary-auto-sync',
    title: 'Character dictionary',
    body: 'Generating character dictionary',
    persistent: true,
  });
  store.upsert({ id: 'one', title: 'One', body: 'First' });
  store.upsert({ id: 'two', title: 'Two', body: 'Second' });
  store.upsert({ id: 'three', title: 'Three', body: 'Third' });

  assert.deepEqual(
    store.visible().map((entry) => entry.id),
    ['character-dictionary-auto-sync', 'two', 'three'],
  );

  store.upsert({
    id: 'character-dictionary-auto-sync',
    title: 'Character dictionary',
    body: 'Ready',
    persistent: false,
  });

  assert.deepEqual(
    store.visible().map((entry) => `${entry.id}:${entry.body}`),
    ['two:Second', 'three:Third', 'character-dictionary-auto-sync:Ready'],
  );
});

test('overlay notification event handler dismisses notifications by id', () => {
  const calls: string[] = [];

  handleOverlayNotificationEvent(
    {
      show: (payload) => {
        calls.push(`show:${payload.id ?? ''}:${payload.title}`);
        return payload.id ?? '';
      },
      remove: (id) => {
        calls.push(`remove:${id}`);
      },
    },
    { id: 'overlay-loading-status', dismiss: true },
  );

  assert.deepEqual(calls, ['remove:overlay-loading-status']);
});

test('overlay notification renderer shows thumbnail image from payload', () => {
  withRenderer({}, ({ renderer, stack }) => {
    renderer.show({
      title: 'Anki Card Updated',
      body: 'Updated card: 食べる',
      image: 'file:///tmp/subminer-notification-icon.png',
      variant: 'success',
      persistent: true,
    });

    const image = requireChild(stack.children[0], 'overlay-notification-image');
    assert.equal(image.tagName, 'IMG');
    assert.equal(image.src, 'file:///tmp/subminer-notification-icon.png');
    assert.equal(image.alt, '');
  });
});

test('overlay notification action buttons send action ids', () => {
  const sentActions: Array<{ notificationId: string; actionId: string; noteId?: number }> = [];
  const electronAPI = {
    sendOverlayNotificationAction: (
      notificationId: string,
      actionId: string,
      options?: { noteId?: number },
    ) => {
      sentActions.push({ notificationId, actionId, noteId: options?.noteId });
    },
  };

  withRenderer({ electronAPI }, ({ renderer, stack }) => {
    renderer.show({
      id: 'subminer-update-available',
      title: 'SubMiner update available',
      body: 'SubMiner v0.15.0 is available',
      persistent: true,
      actions: [{ id: 'open-anki-card', label: 'Open in Anki', noteId: 42 }],
    });

    requireChild(stack.children[0], 'overlay-notification-action').dispatchEventType('click');

    assert.deepEqual(sentActions, [
      { notificationId: 'subminer-update-available', actionId: 'open-anki-card', noteId: 42 },
    ]);
  });
});

test('overlay notification keepOpen actions leave the card on screen', () => {
  const sentActions: string[] = [];
  const electronAPI = {
    sendOverlayNotificationAction: (_notificationId: string, actionId: string) => {
      sentActions.push(actionId);
    },
  };

  withRenderer({ electronAPI }, ({ renderer, stack }) => {
    renderer.show({
      id: 'subminer-update-available',
      title: 'SubMiner update available',
      body: 'SubMiner v0.15.0 is available',
      persistent: true,
      actions: [
        { id: 'install-update', label: 'Update' },
        { id: 'view-changelog', label: "What's New", keepOpen: true },
      ],
    });

    const card = requireChild(stack.children[0]);
    const buttons: FakeElement[] = [];
    const collect = (node: FakeElement): void => {
      if (node.className === 'overlay-notification-action') buttons.push(node);
      for (const child of node.children) collect(child);
    };
    collect(card);
    assert.equal(buttons.length, 2);

    // "What's New" opens the changelog but must not drop the Update affordance.
    buttons[1]?.dispatchEventType('click');
    assert.deepEqual(sentActions, ['view-changelog']);
    assert.equal(card.classList.contains('leaving'), false);

    buttons[0]?.dispatchEventType('click');
    assert.deepEqual(sentActions, ['view-changelog', 'install-update']);
    assert.equal(card.classList.contains('leaving'), true);
  });
});

test('overlay notification renderer updates same-id progress without replacing the spinner', () => {
  withRenderer({}, ({ renderer, stack }) => {
    renderer.show({
      id: 'subsync-status',
      title: 'Subsync',
      body: 'Subsync: syncing |',
      variant: 'progress',
      persistent: true,
    });

    const card = requireChild(stack.children[0]);
    assert.equal(stack.appendCalls, 1);
    assert.equal(card.classList.contains('entering'), true);
    const spinner = requireChild(card, 'overlay-notification-icon');
    const cardReplacements = card.replaceChildrenCalls;

    renderer.show({
      id: 'subsync-status',
      title: 'Subsync',
      body: 'Subsync: syncing /',
      variant: 'progress',
      persistent: true,
    });

    assert.equal(stack.children.length, 1);
    assert.equal(stack.children[0], card);
    assert.equal(stack.appendCalls, 1);
    assert.equal(card.replaceChildrenCalls, cardReplacements);
    assert.equal(findChildByClass(card, 'overlay-notification-icon'), spinner);
    assert.equal(
      findChildByClass(card, 'overlay-notification-body')?.textContent,
      'Subsync: syncing /',
    );

    card.dispatchEventType('animationend', { animationName: 'overlay-notification-enter-right' });
    assert.equal(card.classList.contains('entering'), false);
  });
});

test('overlay notification renderer auto-dismisses same-id terminal update after persistent progress', () => {
  let nextTimerId = 1;
  const timers = new Map<number, { callback: () => void; delayMs: number }>();
  const fakeTimers: Partial<FakeWindow> = {
    clearTimeout: (id) => {
      timers.delete(id);
    },
    setTimeout: (callback, delayMs) => {
      const id = nextTimerId;
      nextTimerId += 1;
      timers.set(id, { callback, delayMs });
      return id;
    },
  };

  withRenderer(fakeTimers, ({ renderer, stack }) => {
    renderer.show({
      id: 'youtube-subtitles-status',
      title: 'YouTube subtitles',
      body: 'Downloading subtitles...',
      variant: 'progress',
      persistent: true,
    });
    const card = requireChild(stack.children[0]);

    renderer.show({
      id: 'youtube-subtitles-status',
      title: 'YouTube subtitles',
      body: 'Subtitles loaded.',
      variant: 'success',
      persistent: false,
    });

    const autoDismissTimer = [...timers.values()].find((timer) => timer.delayMs === 3000);
    assert.ok(autoDismissTimer, 'Expected terminal update to schedule an auto-dismiss timer.');
    assert.equal(stack.children[0], card);
    assert.equal(
      findChildByClass(card, 'overlay-notification-body')?.textContent,
      'Subtitles loaded.',
    );
    assert.equal(card.classList.contains('success'), true);

    autoDismissTimer.callback();

    assert.equal(card.classList.contains('leaving'), true);
  });
});

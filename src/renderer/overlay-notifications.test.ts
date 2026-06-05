import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import test from 'node:test';

import {
  createOverlayNotificationRenderer,
  createOverlayNotificationStore,
  handleOverlayNotificationEvent,
  overlayNotificationPositionClass,
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
  append: (...children: FakeElement[]) => void;
  replaceChildren: (...children: FakeElement[]) => void;
  setAttribute: (name: string, value: string) => void;
  getAttribute: (name: string) => string | null;
  addEventListener: (type: string, listener: (event?: unknown) => void) => void;
};

function createFakeElement(tagName = 'div'): FakeElement {
  const attributes = new Map<string, string>();
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
    append: (...children) => {
      element.children.push(...children);
    },
    replaceChildren: (...children) => {
      element.children = [...children];
    },
    setAttribute: (name, value) => {
      attributes.set(name, value);
    },
    getAttribute: (name) => attributes.get(name) ?? null,
    addEventListener: () => undefined,
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

const overlayNotificationCss = readFileSync(
  path.join(process.cwd(), 'src/renderer/style.css'),
  'utf8',
);

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

test('overlay notification positions map to stack alignment classes', () => {
  assert.equal(overlayNotificationPositionClass(undefined), 'position-top-right');
  assert.equal(overlayNotificationPositionClass('top-left'), 'position-top-left');
  assert.equal(overlayNotificationPositionClass('top'), 'position-top');
  assert.equal(overlayNotificationPositionClass('top-right'), 'position-top-right');
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
  const originalDocument = Object.getOwnPropertyDescriptor(globalThis, 'document');
  const stack = createFakeElement();

  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    writable: true,
    value: {
      createElement: (tagName: string) => createFakeElement(tagName),
    },
  });

  try {
    const renderer = createOverlayNotificationRenderer({
      dom: {
        overlayNotificationStack: stack,
      },
      state: {
        isOverOverlayNotification: false,
      },
    } as never);

    renderer.show({
      title: 'Anki Card Updated',
      body: 'Updated card: 食べる',
      image: 'file:///tmp/subminer-notification-icon.png',
      variant: 'success',
      persistent: true,
    });

    const card = stack.children[0];
    if (!card) {
      assert.fail('Expected overlay notification card.');
    }
    const image = findChildByClass(card, 'overlay-notification-image');
    if (!image) {
      assert.fail('Expected overlay notification image.');
    }

    assert.equal(image.tagName, 'IMG');
    assert.equal(image.src, 'file:///tmp/subminer-notification-icon.png');
    assert.equal(image.alt, '');
  } finally {
    if (originalDocument) {
      Object.defineProperty(globalThis, 'document', originalDocument);
    } else {
      delete (globalThis as { document?: unknown }).document;
    }
  }
});

test('overlay notification cards use larger display dimensions', () => {
  assert.match(
    overlayNotificationCss,
    /\.overlay-notification-stack\s*\{[^}]*width:\s*min\(420px,\s*calc\(100vw - 32px\)\);/s,
  );
  assert.match(overlayNotificationCss, /\.overlay-notification-card\s*\{[^}]*min-height:\s*76px;/s);
  assert.match(
    overlayNotificationCss,
    /\.overlay-notification-card\.has-image\s*\{[^}]*min-height:\s*88px;/s,
  );
  // The has-image card reserves a real grid track for the thumbnail so it
  // cannot overlap the text, and the image shrinks to fit within that track.
  assert.match(
    overlayNotificationCss,
    /\.overlay-notification-card\.has-image\s*\{[^}]*grid-template-columns:\s*minmax\(0,\s*100px\)\s+minmax\(0,\s*1fr\)\s+22px;/s,
  );
  assert.match(overlayNotificationCss, /\.overlay-notification-image\s*\{[^}]*max-width:\s*100px;/s);
  assert.match(
    overlayNotificationCss,
    /\.overlay-notification-image\s*\{[^}]*aspect-ratio:\s*100 \/ 56;/s,
  );
});

import assert from 'node:assert/strict';
import test from 'node:test';

import { createRendererState } from '../state.js';
import { createYoutubeTrackPickerModal } from './youtube-track-picker.js';

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
  };
}

function createFakeElement() {
  const attributes = new Map<string, string>();
  return {
    textContent: '',
    innerHTML: '',
    value: '',
    disabled: false,
    children: [] as any[],
    style: {} as Record<string, string>,
    classList: createClassList(['hidden']),
    listeners: new Map<string, Array<(event?: any) => void>>(),
    appendChild(child: any) {
      this.children.push(child);
      return child;
    },
    append(...children: any[]) {
      this.children.push(...children);
    },
    replaceChildren(...children: any[]) {
      this.children = [...children];
    },
    addEventListener(type: string, listener: (event?: any) => void) {
      const existing = this.listeners.get(type) ?? [];
      existing.push(listener);
      this.listeners.set(type, existing);
    },
    setAttribute(name: string, value: string) {
      attributes.set(name, value);
    },
    getAttribute(name: string) {
      return attributes.get(name) ?? null;
    },
    focus: () => {},
  };
}

function createYoutubePickerDomFixture() {
  return {
    overlay: {
      classList: createClassList(),
      focus: () => {},
    },
    youtubePickerModal: createFakeElement(),
    youtubePickerTitle: createFakeElement(),
    youtubePickerPrimarySelect: createFakeElement(),
    youtubePickerSecondarySelect: createFakeElement(),
    youtubePickerTracks: createFakeElement(),
    youtubePickerStatus: createFakeElement(),
    youtubePickerContinueButton: createFakeElement(),
    youtubePickerCloseButton: createFakeElement(),
  };
}

type YoutubePickerTestWindow = {
  dispatchEvent: (event: Event & { detail?: unknown }) => boolean;
  focus: () => void;
  electronAPI: Record<string, unknown>;
};

function restoreGlobalProp<K extends keyof typeof globalThis>(
  key: K,
  originalValue: (typeof globalThis)[K],
  hadOwnProperty: boolean,
) {
  if (hadOwnProperty) {
    Object.defineProperty(globalThis, key, {
      configurable: true,
      value: originalValue,
      writable: true,
    });
    return;
  }
  Reflect.deleteProperty(globalThis, key);
}

function restoreGlobalDescriptor<K extends keyof typeof globalThis>(
  key: K,
  descriptor: PropertyDescriptor | undefined,
) {
  if (descriptor) {
    Object.defineProperty(globalThis, key, descriptor);
    return;
  }
  Reflect.deleteProperty(globalThis, key);
}

function setupYoutubePickerTestEnv(options?: {
  windowValue?: YoutubePickerTestWindow;
  customEventValue?: unknown;
  now?: () => number;
}) {
  const hadWindow = Object.prototype.hasOwnProperty.call(globalThis, 'window');
  const hadDocument = Object.prototype.hasOwnProperty.call(globalThis, 'document');
  const hadCustomEvent = Object.prototype.hasOwnProperty.call(globalThis, 'CustomEvent');
  const originalWindow = globalThis.window;
  const originalDocument = globalThis.document;
  const originalCustomEvent = globalThis.CustomEvent;
  const originalDateNow = Date.now;

  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createFakeElement(),
    },
    writable: true,
  });

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value:
      options?.windowValue ??
      ({
        dispatchEvent: (_event) => true,
        focus: () => {},
        electronAPI: {
          notifyOverlayModalOpened: () => {},
          notifyOverlayModalClosed: () => {},
          youtubePickerResolve: async () => ({ ok: true, message: '' }),
          setIgnoreMouseEvents: () => {},
        },
      } satisfies YoutubePickerTestWindow),
    writable: true,
  });

  if (options?.customEventValue) {
    Object.defineProperty(globalThis, 'CustomEvent', {
      configurable: true,
      value: options.customEventValue,
      writable: true,
    });
  }

  if (options?.now) {
    Date.now = options.now;
  }

  return {
    restore() {
      Date.now = originalDateNow;
      restoreGlobalProp('window', originalWindow, hadWindow);
      restoreGlobalProp('document', originalDocument, hadDocument);
      restoreGlobalProp('CustomEvent', originalCustomEvent, hadCustomEvent);
    },
  };
}

test('youtube picker test env restore deletes injected globals that were originally absent', () => {
  const previousWindowDescriptor = Object.getOwnPropertyDescriptor(globalThis, 'window');
  const previousDocumentDescriptor = Object.getOwnPropertyDescriptor(globalThis, 'document');

  try {
    Reflect.deleteProperty(globalThis, 'window');
    Reflect.deleteProperty(globalThis, 'document');
    assert.equal(Object.prototype.hasOwnProperty.call(globalThis, 'window'), false);
    assert.equal(Object.prototype.hasOwnProperty.call(globalThis, 'document'), false);

    const env = setupYoutubePickerTestEnv();

    assert.equal(Object.prototype.hasOwnProperty.call(globalThis, 'window'), true);
    assert.equal(Object.prototype.hasOwnProperty.call(globalThis, 'document'), true);

    env.restore();

    assert.equal(Object.prototype.hasOwnProperty.call(globalThis, 'window'), false);
    assert.equal(Object.prototype.hasOwnProperty.call(globalThis, 'document'), false);
    assert.equal(typeof globalThis.window, 'undefined');
    assert.equal(typeof globalThis.document, 'undefined');
  } finally {
    restoreGlobalDescriptor('window', previousWindowDescriptor);
    restoreGlobalDescriptor('document', previousDocumentDescriptor);
  }
});

test('youtube track picker close restores focus and mouse-ignore state', () => {
  const overlayFocusCalls: number[] = [];
  const windowFocusCalls: number[] = [];
  const focusMainWindowCalls: number[] = [];
  const ignoreCalls: Array<{ ignore: boolean; forward?: boolean }> = [];
  const notifications: string[] = [];
  const frontendCommands: unknown[] = [];
  const syncCalls: string[] = [];

  class TestCustomEvent extends Event {
    detail: unknown;

    constructor(type: string, init?: { detail?: unknown }) {
      super(type);
      this.detail = init?.detail;
    }
  }

  const env = setupYoutubePickerTestEnv({
    windowValue: {
      dispatchEvent: (event: Event & { detail?: unknown }) => {
        frontendCommands.push(event.detail ?? null);
        return true;
      },
      focus: () => {
        windowFocusCalls.push(1);
      },
      electronAPI: {
        notifyOverlayModalOpened: () => {},
        notifyOverlayModalClosed: (modal: string) => {
          notifications.push(modal);
        },
        focusMainWindow: async () => {
          focusMainWindowCalls.push(1);
        },
        youtubePickerResolve: async () => ({ ok: true, message: '' }),
        setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
          ignoreCalls.push({ ignore, forward: options?.forward });
        },
      },
    } satisfies YoutubePickerTestWindow,
    customEventValue: TestCustomEvent,
  });

  try {
    const state = createRendererState();
    const dom = createYoutubePickerDomFixture();
    dom.overlay.focus = () => {
      overlayFocusCalls.push(1);
    };
    const { overlay } = dom;

    const modal = createYoutubeTrackPickerModal(
      {
        state,
        dom,
        platform: {
          shouldToggleMouseIgnore: true,
        },
      } as never,
      {
        modalStateReader: { isAnyModalOpen: () => false },
        restorePointerInteractionState: () => {
          syncCalls.push('restore-pointer');
        },
        syncSettingsModalSubtitleSuppression: () => {
          syncCalls.push('sync');
        },
      },
    );

    modal.openYoutubePickerModal({
      sessionId: 'yt-1',
      url: 'https://example.com',
      tracks: [],
      defaultPrimaryTrackId: null,
      defaultSecondaryTrackId: null,
      hasTracks: false,
    });
    modal.closeYoutubePickerModal();

    assert.equal(state.youtubePickerModalOpen, false);
    assert.deepEqual(syncCalls, ['sync', 'sync', 'restore-pointer']);
    assert.deepEqual(notifications, ['youtube-track-picker']);
    assert.deepEqual(frontendCommands, [{ type: 'refreshOptions' }]);
    assert.equal(overlay.classList.contains('interactive'), false);
    assert.equal(focusMainWindowCalls.length > 0, true);
    assert.equal(overlayFocusCalls.length > 0, true);
    assert.equal(windowFocusCalls.length > 0, true);
    assert.deepEqual(ignoreCalls, [{ ignore: true, forward: true }]);
  } finally {
    env.restore();
  }
});

test('youtube track picker re-acknowledges repeated open requests', () => {
  const openedNotifications: string[] = [];
  const env = setupYoutubePickerTestEnv({
    windowValue: {
      dispatchEvent: (_event) => true,
      focus: () => {},
      electronAPI: {
        notifyOverlayModalOpened: (modal: string) => {
          openedNotifications.push(modal);
        },
        notifyOverlayModalClosed: () => {},
        youtubePickerResolve: async () => ({ ok: true, message: '' }),
        setIgnoreMouseEvents: () => {},
      },
    } satisfies YoutubePickerTestWindow,
  });

  try {
    const state = createRendererState();
    const dom = createYoutubePickerDomFixture();

    const modal = createYoutubeTrackPickerModal(
      {
        state,
        dom,
        platform: {
          shouldToggleMouseIgnore: false,
        },
      } as never,
      {
        modalStateReader: { isAnyModalOpen: () => true },
        restorePointerInteractionState: () => {},
        syncSettingsModalSubtitleSuppression: () => {},
      },
    );

    modal.openYoutubePickerModal({
      sessionId: 'yt-1',
      url: 'https://example.com/one',
      tracks: [],
      defaultPrimaryTrackId: null,
      defaultSecondaryTrackId: null,
      hasTracks: false,
    });
    modal.openYoutubePickerModal({
      sessionId: 'yt-2',
      url: 'https://example.com/two',
      tracks: [],
      defaultPrimaryTrackId: null,
      defaultSecondaryTrackId: null,
      hasTracks: false,
    });

    assert.deepEqual(openedNotifications, ['youtube-track-picker', 'youtube-track-picker']);
    assert.equal(state.youtubePickerPayload?.sessionId, 'yt-2');
  } finally {
    env.restore();
  }
});

test('youtube track picker surfaces rejected resolve calls as modal status', async () => {
  const env = setupYoutubePickerTestEnv({
    windowValue: {
      dispatchEvent: (_event) => true,
      focus: () => {},
      electronAPI: {
        notifyOverlayModalOpened: () => {},
        notifyOverlayModalClosed: () => {},
        youtubePickerResolve: async () => {
          throw new Error('resolve failed');
        },
        setIgnoreMouseEvents: () => {},
      },
    } satisfies YoutubePickerTestWindow,
  });

  try {
    const state = createRendererState();
    const dom = createYoutubePickerDomFixture();

    const modal = createYoutubeTrackPickerModal(
      {
        state,
        dom,
        platform: {
          shouldToggleMouseIgnore: false,
        },
      } as never,
      {
        modalStateReader: { isAnyModalOpen: () => true },
        restorePointerInteractionState: () => {},
        syncSettingsModalSubtitleSuppression: () => {},
      },
    );

    modal.openYoutubePickerModal({
      sessionId: 'yt-1',
      url: 'https://example.com',
      tracks: [
        {
          id: 'auto:ja-orig',
          language: 'ja',
          sourceLanguage: 'ja-orig',
          kind: 'auto',
          label: 'Japanese (auto)',
        },
      ],
      defaultPrimaryTrackId: 'auto:ja-orig',
      defaultSecondaryTrackId: null,
      hasTracks: true,
    });
    modal.wireDomEvents();

    const listeners = dom.youtubePickerContinueButton.listeners.get('click') ?? [];
    await Promise.all(listeners.map((listener) => listener()));

    assert.equal(state.youtubePickerModalOpen, true);
    assert.equal(dom.youtubePickerStatus.textContent, 'resolve failed');
  } finally {
    env.restore();
  }
});

test('youtube track picker ignores duplicate resolve submissions while request is in flight', async () => {
  const resolveCalls: Array<{
    sessionId: string;
    action: string;
    primaryTrackId: string | null;
    secondaryTrackId: string | null;
  }> = [];
  let releaseResolve: (() => void) | null = null;
  const env = setupYoutubePickerTestEnv({
    windowValue: {
      dispatchEvent: (_event) => true,
      focus: () => {},
      electronAPI: {
        notifyOverlayModalOpened: () => {},
        notifyOverlayModalClosed: () => {},
        youtubePickerResolve: async (payload: {
          sessionId: string;
          action: string;
          primaryTrackId: string | null;
          secondaryTrackId: string | null;
        }) => {
          resolveCalls.push(payload);
          await new Promise<void>((resolve) => {
            releaseResolve = resolve;
          });
          return { ok: true, message: '' };
        },
        setIgnoreMouseEvents: () => {},
      },
    } satisfies YoutubePickerTestWindow,
  });

  try {
    const state = createRendererState();
    const dom = createYoutubePickerDomFixture();

    const modal = createYoutubeTrackPickerModal(
      {
        state,
        dom,
        platform: {
          shouldToggleMouseIgnore: false,
        },
      } as never,
      {
        modalStateReader: { isAnyModalOpen: () => true },
        restorePointerInteractionState: () => {},
        syncSettingsModalSubtitleSuppression: () => {},
      },
    );

    modal.openYoutubePickerModal({
      sessionId: 'yt-1',
      url: 'https://example.com',
      tracks: [
        {
          id: 'auto:ja-orig',
          language: 'ja',
          sourceLanguage: 'ja-orig',
          kind: 'auto',
          label: 'Japanese (auto)',
        },
      ],
      defaultPrimaryTrackId: 'auto:ja-orig',
      defaultSecondaryTrackId: null,
      hasTracks: true,
    });
    modal.wireDomEvents();

    const listeners = dom.youtubePickerContinueButton.listeners.get('click') ?? [];
    const first = listeners[0]?.();
    const second = listeners[0]?.();
    await Promise.resolve();

    assert.equal(resolveCalls.length, 1);
    assert.equal(dom.youtubePickerPrimarySelect.disabled, true);
    assert.equal(dom.youtubePickerSecondarySelect.disabled, true);
    assert.equal(dom.youtubePickerContinueButton.disabled, true);
    assert.equal(dom.youtubePickerCloseButton.disabled, true);

    assert.ok(releaseResolve);
    const release = releaseResolve as () => void;
    release();
    await Promise.all([first, second]);

    assert.equal(dom.youtubePickerPrimarySelect.disabled, false);
    assert.equal(dom.youtubePickerSecondarySelect.disabled, false);
    assert.equal(dom.youtubePickerContinueButton.disabled, false);
    assert.equal(dom.youtubePickerCloseButton.disabled, false);
  } finally {
    env.restore();
  }
});

test('youtube track picker keeps no-track controls disabled after a rejected continue request', async () => {
  const env = setupYoutubePickerTestEnv({
    windowValue: {
      dispatchEvent: (_event) => true,
      focus: () => {},
      electronAPI: {
        notifyOverlayModalOpened: () => {},
        notifyOverlayModalClosed: () => {},
        youtubePickerResolve: async () => ({ ok: false, message: 'still no tracks' }),
        setIgnoreMouseEvents: () => {},
      },
    } satisfies YoutubePickerTestWindow,
  });

  try {
    const state = createRendererState();
    const dom = createYoutubePickerDomFixture();

    const modal = createYoutubeTrackPickerModal(
      {
        state,
        dom,
        platform: {
          shouldToggleMouseIgnore: false,
        },
      } as never,
      {
        modalStateReader: { isAnyModalOpen: () => true },
        restorePointerInteractionState: () => {},
        syncSettingsModalSubtitleSuppression: () => {},
      },
    );

    modal.openYoutubePickerModal({
      sessionId: 'yt-1',
      url: 'https://example.com',
      tracks: [],
      defaultPrimaryTrackId: null,
      defaultSecondaryTrackId: null,
      hasTracks: false,
    });
    modal.wireDomEvents();

    const listeners = dom.youtubePickerContinueButton.listeners.get('click') ?? [];
    await Promise.all(listeners.map((listener) => listener()));

    assert.equal(dom.youtubePickerPrimarySelect.disabled, true);
    assert.equal(dom.youtubePickerSecondarySelect.disabled, true);
    assert.equal(dom.youtubePickerContinueButton.disabled, true);
    assert.equal(dom.youtubePickerCloseButton.disabled, true);
    assert.equal(dom.youtubePickerStatus.textContent, 'still no tracks');
  } finally {
    env.restore();
  }
});

test('youtube track picker only consumes handled keys', async () => {
  const env = setupYoutubePickerTestEnv();

  try {
    const state = createRendererState();
    const dom = createYoutubePickerDomFixture();

    const modal = createYoutubeTrackPickerModal(
      {
        state,
        dom,
        platform: {
          shouldToggleMouseIgnore: false,
        },
      } as never,
      {
        modalStateReader: { isAnyModalOpen: () => true },
        restorePointerInteractionState: () => {},
        syncSettingsModalSubtitleSuppression: () => {},
      },
    );

    modal.openYoutubePickerModal({
      sessionId: 'yt-1',
      url: 'https://example.com',
      tracks: [],
      defaultPrimaryTrackId: null,
      defaultSecondaryTrackId: null,
      hasTracks: false,
    });

    assert.equal(
      modal.handleYoutubePickerKeydown({
        key: ' ',
        preventDefault: () => {},
      } as KeyboardEvent),
      false,
    );
    assert.equal(
      modal.handleYoutubePickerKeydown({
        key: 'Escape',
        preventDefault: () => {},
      } as KeyboardEvent),
      true,
    );
    await Promise.resolve();
  } finally {
    env.restore();
  }
});

test('youtube track picker ignores immediate Enter after open before allowing keyboard submit', async () => {
  const resolveCalls: Array<{
    sessionId: string;
    action: string;
    primaryTrackId: string | null;
    secondaryTrackId: string | null;
  }> = [];
  let now = 10_000;
  const env = setupYoutubePickerTestEnv({
    now: () => now,
    windowValue: {
      dispatchEvent: (_event) => true,
      focus: () => {},
      electronAPI: {
        notifyOverlayModalOpened: () => {},
        notifyOverlayModalClosed: () => {},
        youtubePickerResolve: async (payload: {
          sessionId: string;
          action: string;
          primaryTrackId: string | null;
          secondaryTrackId: string | null;
        }) => {
          resolveCalls.push(payload);
          return { ok: true, message: '' };
        },
        setIgnoreMouseEvents: () => {},
      },
    } satisfies YoutubePickerTestWindow,
  });

  try {
    const state = createRendererState();
    const dom = createYoutubePickerDomFixture();

    const modal = createYoutubeTrackPickerModal(
      {
        state,
        dom,
        platform: {
          shouldToggleMouseIgnore: false,
        },
      } as never,
      {
        modalStateReader: { isAnyModalOpen: () => true },
        restorePointerInteractionState: () => {},
        syncSettingsModalSubtitleSuppression: () => {},
      },
    );

    modal.openYoutubePickerModal({
      sessionId: 'yt-1',
      url: 'https://example.com',
      tracks: [
        {
          id: 'auto:ja-orig',
          language: 'ja',
          sourceLanguage: 'ja-orig',
          kind: 'auto',
          label: 'Japanese (auto)',
        },
      ],
      defaultPrimaryTrackId: 'auto:ja-orig',
      defaultSecondaryTrackId: null,
      hasTracks: true,
    });

    assert.equal(
      modal.handleYoutubePickerKeydown({
        key: 'Enter',
        preventDefault: () => {},
      } as KeyboardEvent),
      true,
    );
    await Promise.resolve();
    assert.deepEqual(resolveCalls, []);
    assert.equal(state.youtubePickerModalOpen, true);

    now += 250;
    assert.equal(
      modal.handleYoutubePickerKeydown({
        key: 'Enter',
        preventDefault: () => {},
      } as KeyboardEvent),
      true,
    );
    await Promise.resolve();
    assert.deepEqual(resolveCalls, [
      {
        sessionId: 'yt-1',
        action: 'use-selected',
        primaryTrackId: 'auto:ja-orig',
        secondaryTrackId: null,
      },
    ]);
  } finally {
    env.restore();
  }
});

test('youtube track picker uses track list as the source of truth for available selections', async () => {
  const resolveCalls: Array<{
    sessionId: string;
    action: string;
    primaryTrackId: string | null;
    secondaryTrackId: string | null;
  }> = [];
  const env = setupYoutubePickerTestEnv({
    windowValue: {
      dispatchEvent: (_event) => true,
      focus: () => {},
      electronAPI: {
        notifyOverlayModalOpened: () => {},
        notifyOverlayModalClosed: () => {},
        youtubePickerResolve: async (payload: {
          sessionId: string;
          action: string;
          primaryTrackId: string | null;
          secondaryTrackId: string | null;
        }) => {
          resolveCalls.push(payload);
          return { ok: true, message: '' };
        },
        setIgnoreMouseEvents: () => {},
      },
    } satisfies YoutubePickerTestWindow,
  });

  try {
    const state = createRendererState();
    const dom = createYoutubePickerDomFixture();

    const modal = createYoutubeTrackPickerModal(
      {
        state,
        dom,
        platform: {
          shouldToggleMouseIgnore: false,
        },
      } as never,
      {
        modalStateReader: { isAnyModalOpen: () => true },
        restorePointerInteractionState: () => {},
        syncSettingsModalSubtitleSuppression: () => {},
      },
    );

    modal.openYoutubePickerModal({
      sessionId: 'yt-1',
      url: 'https://example.com',
      tracks: [
        {
          id: 'manual:ja',
          language: 'ja',
          sourceLanguage: 'ja',
          kind: 'manual',
          label: 'Japanese',
        },
      ],
      defaultPrimaryTrackId: 'manual:ja',
      defaultSecondaryTrackId: null,
      hasTracks: false,
    });
    modal.wireDomEvents();

    const listeners = dom.youtubePickerContinueButton.listeners.get('click') ?? [];
    await Promise.all(listeners.map((listener) => listener()));

    assert.deepEqual(resolveCalls, [
      {
        sessionId: 'yt-1',
        action: 'use-selected',
        primaryTrackId: 'manual:ja',
        secondaryTrackId: null,
      },
    ]);
  } finally {
    env.restore();
  }
});

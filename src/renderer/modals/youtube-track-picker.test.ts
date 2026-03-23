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

test('youtube track picker close restores focus and mouse-ignore state', () => {
  const overlayFocusCalls: number[] = [];
  const windowFocusCalls: number[] = [];
  const ignoreCalls: Array<{ ignore: boolean; forward?: boolean }> = [];
  const notifications: string[] = [];
  const frontendCommands: unknown[] = [];
  const syncCalls: string[] = [];
  const originalWindow = globalThis.window;
  const originalDocument = globalThis.document;
  const originalCustomEvent = globalThis.CustomEvent;

  class TestCustomEvent extends Event {
    detail: unknown;

    constructor(type: string, init?: { detail?: unknown }) {
      super(type);
      this.detail = init?.detail;
    }
  }

  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createFakeElement(),
    },
  });

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
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
        youtubePickerResolve: async () => ({ ok: true, message: '' }),
        setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
          ignoreCalls.push({ ignore, forward: options?.forward });
        },
      },
    },
  });

  Object.defineProperty(globalThis, 'CustomEvent', {
    configurable: true,
    value: TestCustomEvent,
  });

  try {
    const state = createRendererState();
    const overlay = {
      classList: createClassList(),
      focus: () => {
        overlayFocusCalls.push(1);
      },
    };
    const dom = {
      overlay,
      youtubePickerModal: createFakeElement(),
      youtubePickerTitle: createFakeElement(),
      youtubePickerPrimarySelect: createFakeElement(),
      youtubePickerSecondarySelect: createFakeElement(),
      youtubePickerTracks: createFakeElement(),
      youtubePickerStatus: createFakeElement(),
      youtubePickerContinueButton: createFakeElement(),
      youtubePickerCloseButton: createFakeElement(),
    };

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
      mode: 'download',
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
    assert.equal(overlayFocusCalls.length > 0, true);
    assert.equal(windowFocusCalls.length > 0, true);
    assert.deepEqual(ignoreCalls, [{ ignore: true, forward: true }]);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: originalWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: originalDocument });
    Object.defineProperty(globalThis, 'CustomEvent', {
      configurable: true,
      value: originalCustomEvent,
    });
  }
});

import assert from 'node:assert/strict';
import test from 'node:test';
import { createAnimeBrowserModal } from './anime-browser';

function createClassList(initial: string[] = []) {
  const classes = new Set(initial);
  return {
    add: (...names: string[]) => names.forEach((name) => classes.add(name)),
    remove: (...names: string[]) => names.forEach((name) => classes.delete(name)),
    contains: (name: string) => classes.has(name),
  };
}

test('embedded Anime Browser closes only for its own close message', () => {
  const previousWindow = Object.getOwnPropertyDescriptor(globalThis, 'window');
  const messages: Array<(event: MessageEvent) => void> = [];
  const notifications: string[] = [];
  const frameWindow = {};
  const closeListeners: Array<() => void> = [];
  const removedCloseListeners: Array<() => void> = [];
  const overlay = { classList: createClassList() };
  const modalElement = {
    classList: createClassList(['hidden']),
    attributes: new Map<string, string>(),
    setAttribute(name: string, value: string) {
      this.attributes.set(name, value);
    },
  };
  const frame = {
    src: '',
    dataset: { src: '../animeui/index.html?embedded=overlay-modal' },
    contentWindow: frameWindow,
  };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        notifyOverlayModalOpened: (name: string) => notifications.push(`open:${name}`),
        notifyOverlayModalClosed: (name: string) => notifications.push(`close:${name}`),
      },
      addEventListener: (name: string, listener: (event: MessageEvent) => void) => {
        if (name === 'message') messages.push(listener);
      },
      removeEventListener: () => undefined,
    },
  });

  try {
    const state = { animeBrowserModalOpen: false, isOverSubtitle: false };
    const modal = createAnimeBrowserModal(
      {
        state,
        dom: {
          overlay,
          animeBrowserModal: modalElement,
          animeBrowserFrame: frame,
          animeBrowserClose: {
            addEventListener: (_name: string, listener: () => void) =>
              closeListeners.push(listener),
            removeEventListener: (_name: string, listener: () => void) =>
              removedCloseListeners.push(listener),
          },
        },
      } as never,
      {
        modalStateReader: { isAnyModalOpen: () => state.animeBrowserModalOpen },
        syncSettingsModalSubtitleSuppression: () => undefined,
      },
    );

    modal.wireDomEvents();
    modal.openAnimeBrowserModal();
    assert.equal(state.animeBrowserModalOpen, true);
    assert.equal(frame.src, frame.dataset.src);
    assert.deepEqual(notifications, ['open:anime-browser']);

    messages[0]?.({ source: {}, data: 'subminer:anime-browser-close' } as MessageEvent);
    messages[0]?.({ source: frameWindow, data: 'not-the-close-message' } as MessageEvent);
    assert.equal(state.animeBrowserModalOpen, true);

    messages[0]?.({ source: frameWindow, data: 'subminer:anime-browser-close' } as MessageEvent);
    assert.equal(state.animeBrowserModalOpen, false);
    assert.equal(modalElement.classList.contains('hidden'), true);
    assert.deepEqual(notifications, ['open:anime-browser', 'close:anime-browser']);
    modal.disposeDomEvents();
    assert.deepEqual(removedCloseListeners, closeListeners);
  } finally {
    if (previousWindow) Object.defineProperty(globalThis, 'window', previousWindow);
    else Reflect.deleteProperty(globalThis, 'window');
  }
});

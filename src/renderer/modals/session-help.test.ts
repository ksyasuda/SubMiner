import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';

import { SPECIAL_COMMANDS } from '../../config/definitions/shared';
import { createRendererState } from '../state.js';
import {
  createSessionHelpModal,
  describeSessionHelpCommand,
  formatSessionHelpKeybinding,
} from './session-help.js';

test('session help describes sub-seek commands as subtitle-line navigation', () => {
  assert.equal(describeSessionHelpCommand(['sub-seek', 1]), 'Jump to next subtitle');
  assert.equal(describeSessionHelpCommand(['sub-seek', -1]), 'Jump to previous subtitle');
});

test('session help describes subtitle-delay shift special commands separately from sub-seek', () => {
  assert.equal(
    describeSessionHelpCommand([SPECIAL_COMMANDS.SHIFT_SUB_DELAY_TO_NEXT_SUBTITLE_START]),
    'Shift subtitle delay to next cue',
  );
  assert.equal(
    describeSessionHelpCommand([SPECIAL_COMMANDS.SHIFT_SUB_DELAY_TO_PREVIOUS_SUBTITLE_START]),
    'Shift subtitle delay to previous cue',
  );
});

test('session help formats bracket keybindings as physical keys', () => {
  assert.equal(formatSessionHelpKeybinding('Shift+BracketRight'), 'Shift + ]');
  assert.equal(formatSessionHelpKeybinding('Shift+BracketLeft'), 'Shift + [');
});

test('session help imports browser-safe special command constants', () => {
  const source = fs.readFileSync(
    path.join(process.cwd(), 'src', 'renderer', 'modals', 'session-help.ts'),
    'utf8',
  );

  assert.match(source, /from ['"]\.\.\/\.\.\/config\/definitions\/shared['"]/);
  assert.doesNotMatch(source, /from ['"]\.\.\/\.\.\/config\/definitions['"]/);
});

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

function createElementStub() {
  return {
    value: '',
    textContent: '',
    innerHTML: '',
    classList: createClassList(['hidden']),
    setAttribute: () => {},
    addEventListener: () => {},
    removeEventListener: () => {},
    querySelectorAll: () => [],
    focus: () => {},
    select: () => {},
  };
}

test('modal-layer session help does not focus hidden main overlay and still closes', async () => {
  const globals = globalThis as typeof globalThis & {
    window?: unknown;
    document?: unknown;
    HTMLElement?: unknown;
    Element?: unknown;
  };
  const previousWindow = globals.window;
  const previousDocument = globals.document;
  const previousHTMLElement = globals.HTMLElement;
  const previousElement = globals.Element;
  const focusMainWindowCalls: number[] = [];
  const notifications: string[] = [];

  try {
    class TestElement {}
    Object.defineProperty(globalThis, 'HTMLElement', {
      configurable: true,
      writable: true,
      value: TestElement,
    });
    Object.defineProperty(globalThis, 'Element', {
      configurable: true,
      writable: true,
      value: TestElement,
    });
    Object.defineProperty(globalThis, 'window', {
      configurable: true,
      writable: true,
      value: {
        electronAPI: {
          focusMainWindow: async () => {
            focusMainWindowCalls.push(1);
          },
          setIgnoreMouseEvents: () => {},
          notifyOverlayModalClosed: (modal: string) => {
            notifications.push(modal);
          },
          getKeybindings: async () => {
            throw new Error('mpv unavailable');
          },
          getSubtitleStyle: async () => ({}),
          getConfiguredShortcuts: async () => ({}),
        },
        focus: () => {},
        addEventListener: () => {},
        removeEventListener: () => {},
        setTimeout: (callback: () => void) => setTimeout(callback, 0),
      },
    });
    Object.defineProperty(globalThis, 'document', {
      configurable: true,
      writable: true,
      value: {
        activeElement: null,
        addEventListener: () => {},
        removeEventListener: () => {},
      },
    });

    const state = createRendererState();
    const modal = createSessionHelpModal(
      {
        state,
        platform: {
          overlayLayer: 'modal',
          isModalLayer: true,
          isLinuxPlatform: false,
          isMacOSPlatform: false,
          isWindowsPlatform: true,
          shouldToggleMouseIgnore: false,
        },
        dom: {
          overlay: createElementStub(),
          sessionHelpModal: createElementStub(),
          sessionHelpFilter: createElementStub(),
          sessionHelpContent: createElementStub(),
          sessionHelpClose: createElementStub(),
          sessionHelpShortcut: createElementStub(),
          sessionHelpWarning: createElementStub(),
          sessionHelpStatus: createElementStub(),
        },
      } as never,
      {
        modalStateReader: { isAnyModalOpen: () => false },
        syncSettingsModalSubtitleSuppression: () => {},
      },
    );

    modal.openSessionHelpModal({
      bindingKey: 'KeyH',
      fallbackUsed: false,
      fallbackUnavailable: false,
    });
    modal.closeSessionHelpModal();
    await new Promise((resolve) => setTimeout(resolve, 0));

    assert.deepEqual(focusMainWindowCalls, []);
    assert.deepEqual(notifications, ['session-help']);
  } finally {
    Object.defineProperty(globalThis, 'window', {
      configurable: true,
      writable: true,
      value: previousWindow,
    });
    Object.defineProperty(globalThis, 'document', {
      configurable: true,
      writable: true,
      value: previousDocument,
    });
    Object.defineProperty(globalThis, 'HTMLElement', {
      configurable: true,
      writable: true,
      value: previousHTMLElement,
    });
    Object.defineProperty(globalThis, 'Element', {
      configurable: true,
      writable: true,
      value: previousElement,
    });
  }
});

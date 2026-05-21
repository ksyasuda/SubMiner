import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';

import { SPECIAL_COMMANDS } from '../../config/definitions/shared';
import { createRendererState } from '../state.js';
import {
  buildSessionHelpSections,
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

test('session help normalizes configured modifier aliases', () => {
  assert.equal(formatSessionHelpKeybinding('CommandOrControl+KeyS'), 'Cmd/Ctrl + S');
});

test('session help imports browser-safe special command constants', () => {
  const source = fs.readFileSync(
    path.join(process.cwd(), 'src', 'renderer', 'modals', 'session-help-sections.ts'),
    'utf8',
  );

  assert.match(source, /from ['"]\.\.\/\.\.\/config\/definitions\/shared['"]/);
  assert.doesNotMatch(source, /from ['"]\.\.\/\.\.\/config\/definitions['"]/);
});

test('session help builds rows from canonical session bindings and fixed overlay affordances', () => {
  const sections = buildSessionHelpSections({
    sessionBindings: [
      {
        sourcePath: 'stats.toggleKey',
        originalKey: 'Backquote',
        key: { code: 'Backquote', modifiers: [] },
        actionType: 'session-action',
        actionId: 'toggleStatsOverlay',
      },
      {
        sourcePath: 'shortcuts.openSessionHelp',
        originalKey: 'CommandOrControl+Slash',
        key: { code: 'Slash', modifiers: ['ctrl'] },
        actionType: 'session-action',
        actionId: 'openSessionHelp',
      },
      {
        sourcePath: 'shortcuts.toggleSubtitleSidebar',
        originalKey: 'Backslash',
        key: { code: 'Backslash', modifiers: [] },
        actionType: 'session-action',
        actionId: 'toggleSubtitleSidebar',
      },
      {
        sourcePath: 'stats.markWatchedKey',
        originalKey: 'KeyW',
        key: { code: 'KeyW', modifiers: [] },
        actionType: 'session-action',
        actionId: 'markWatched',
      },
      {
        sourcePath: 'keybindings[0].key',
        originalKey: 'Space',
        key: { code: 'Space', modifiers: [] },
        actionType: 'mpv-command',
        command: ['cycle', 'pause'],
      },
    ],
    markWatchedKey: 'KeyW',
    subtitleSidebarToggleKey: 'KeyB',
    subtitleStyle: {},
  });

  const rows = sections.flatMap((section) => section.rows);
  assert.ok(rows.some((row) => row.shortcut === '`' && row.action === 'Toggle stats overlay'));
  assert.ok(rows.some((row) => row.shortcut === 'W' && row.action === 'Mark video watched'));
  assert.equal(rows.filter((row) => row.action === 'Mark video watched').length, 1);
  assert.equal(sections.filter((section) => section.title === 'Stats and progress').length, 1);
  assert.ok(rows.some((row) => row.shortcut === 'B' && row.action === 'Toggle subtitle sidebar'));
  assert.equal(rows.filter((row) => row.action === 'Toggle subtitle sidebar').length, 1);
  assert.ok(rows.some((row) => row.shortcut === 'Ctrl + /' && row.action === 'Open session help'));
  assert.ok(rows.some((row) => row.shortcut === 'Space' && row.action === 'Toggle playback'));
  assert.ok(
    rows.some(
      (row) => row.shortcut === 'V' && row.action === 'Toggle primary subtitle bar visibility',
    ),
  );
  assert.ok(rows.some((row) => row.shortcut === 'Y then D' && row.action === 'Toggle DevTools'));
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
          getSessionBindings: async () => [],
          getSubtitleStyle: async () => ({}),
          getMarkWatchedKey: async () => 'KeyW',
          getSubtitleSidebarSnapshot: async () => ({
            config: { toggleKey: 'Backslash' },
          }),
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

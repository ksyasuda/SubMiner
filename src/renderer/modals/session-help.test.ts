import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';

import { createRendererState } from '../state.js';
import {
  buildSessionHelpSections,
  createSessionHelpModal,
  describeSessionHelpCommand,
  formatSessionHelpKeybinding,
  isKnownWordMaturityLegendEnabled,
} from './session-help.js';
import type { RuntimeOptionId, RuntimeOptionState } from '../../types/runtime-options.js';

test('session help describes sub-seek commands as subtitle-line navigation', () => {
  assert.equal(describeSessionHelpCommand(['sub-seek', 1]), 'Jump to next subtitle');
  assert.equal(describeSessionHelpCommand(['sub-seek', -1]), 'Jump to previous subtitle');
});

test('session help describes native subtitle-delay step commands separately from sub-seek', () => {
  assert.equal(describeSessionHelpCommand(['sub-step', 1]), 'Shift subtitle delay to next cue');
  assert.equal(
    describeSessionHelpCommand(['sub-step', -1]),
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

function booleanRuntimeOption(id: RuntimeOptionId, value: boolean): RuntimeOptionState {
  return {
    id,
    label: id,
    scope: 'subtitle',
    valueType: 'boolean',
    value,
    allowedValues: [true, false],
    requiresRestart: false,
  };
}

test('maturity legend requires both known-word highlighting and maturity coloring', () => {
  const highlightOn = booleanRuntimeOption('subtitle.annotation.knownWords.highlightEnabled', true);
  const highlightOff = booleanRuntimeOption(
    'subtitle.annotation.knownWords.highlightEnabled',
    false,
  );
  const maturityOn = booleanRuntimeOption('subtitle.annotation.knownWords.maturityEnabled', true);
  const maturityOff = booleanRuntimeOption('subtitle.annotation.knownWords.maturityEnabled', false);

  assert.equal(isKnownWordMaturityLegendEnabled([highlightOn, maturityOn]), true);
  assert.equal(isKnownWordMaturityLegendEnabled([highlightOff, maturityOn]), false);
  assert.equal(isKnownWordMaturityLegendEnabled([highlightOn, maturityOff]), false);
  assert.equal(isKnownWordMaturityLegendEnabled([maturityOn]), false);
  assert.equal(isKnownWordMaturityLegendEnabled([]), false);
});

function colorLegendRows(input: Parameters<typeof buildSessionHelpSections>[0]) {
  const sections = buildSessionHelpSections(input);
  return sections.find((section) => section.title === 'Color legend')?.rows ?? [];
}

test('color legend shows the flat known-word color when maturity coloring is off', () => {
  const rows = colorLegendRows({
    sessionBindings: [],
    subtitleStyle: {
      knownWordColor: '#a6da95',
      knownWordMaturityColors: {
        new: '#ee99a0',
        learning: '#b7bdf8',
        young: '#91d7e3',
        mature: '#a6da95',
      },
    },
  });

  assert.deepEqual(
    rows.filter((row) => row.shortcut.startsWith('Known words')),
    [{ shortcut: 'Known words', action: '#a6da95', color: '#a6da95' }],
  );
});

test('color legend swaps in maturity tiers when maturity coloring is on', () => {
  const rows = colorLegendRows({
    sessionBindings: [],
    subtitleStyle: {
      knownWordColor: '#a6da95',
      knownWordMaturityColors: {
        new: '#ee99a0',
        learning: '#b7bdf8',
        young: '#91d7e3',
        mature: '#f0c6c6',
      },
    },
    knownWordMaturityEnabled: true,
  });

  assert.deepEqual(
    rows.filter((row) => row.shortcut.startsWith('Known words')),
    [
      { shortcut: 'Known words (new)', action: '#ee99a0', color: '#ee99a0' },
      { shortcut: 'Known words (learning)', action: '#b7bdf8', color: '#b7bdf8' },
      { shortcut: 'Known words (young)', action: '#91d7e3', color: '#91d7e3' },
      { shortcut: 'Known words (mature)', action: '#f0c6c6', color: '#f0c6c6' },
    ],
  );
  assert.ok(rows.some((row) => row.shortcut === 'N+1 words'));
});

test('color legend falls back to default maturity colors when overrides are invalid', () => {
  const rows = colorLegendRows({
    sessionBindings: [],
    subtitleStyle: {
      knownWordMaturityColors: { new: 'not-a-color', learning: 42 },
    },
    knownWordMaturityEnabled: true,
  });

  assert.deepEqual(
    rows.filter((row) => row.shortcut.startsWith('Known words')).map((row) => row.color),
    ['#ee99a0', '#b7bdf8', '#91d7e3', '#a6da95'],
  );
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
          getRuntimeOptions: async () => [],
        },
        focus: () => {},
        addEventListener: () => {},
        removeEventListener: () => {},
        setTimeout: (callback: () => void) => setTimeout(callback, 0),
        clearTimeout: (id: unknown) => clearTimeout(id as ReturnType<typeof setTimeout>),
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

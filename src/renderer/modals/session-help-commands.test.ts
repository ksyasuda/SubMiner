import assert from 'node:assert/strict';
import test from 'node:test';

import type { CompiledSessionBinding } from '../../types.js';
import { createRendererState } from '../state.js';
import { buildSessionHelpSections, createSessionHelpModal } from './session-help.js';

/** Just enough DOM for the help modal to render rows and route row events. */
class FakeElement {
  children: FakeElement[] = [];
  parent: FakeElement | null = null;
  dataset: Record<string, string> = {};
  style: Record<string, string> = {};
  textContent = '';
  value = '';
  tabIndex = 0;
  type = '';
  private classes = new Set<string>();
  private listeners = new Map<string, Array<(event: unknown) => void>>();

  classList = {
    add: (...tokens: string[]) => tokens.forEach((token) => this.classes.add(token)),
    remove: (...tokens: string[]) => tokens.forEach((token) => this.classes.delete(token)),
    toggle: (token: string, force?: boolean) => {
      const on = force ?? !this.classes.has(token);
      if (on) this.classes.add(token);
      else this.classes.delete(token);
      return on;
    },
    contains: (token: string) => this.classes.has(token),
  };

  set className(value: string) {
    this.classes = new Set(value.split(/\s+/).filter(Boolean));
  }

  get className(): string {
    return [...this.classes].join(' ');
  }

  set innerHTML(_value: string) {
    this.children = [];
  }

  appendChild(child: FakeElement): FakeElement {
    child.parent = this;
    this.children.push(child);
    return child;
  }

  insertBefore(child: FakeElement, ref: FakeElement): FakeElement {
    child.parent = this;
    const index = this.children.indexOf(ref);
    this.children.splice(index < 0 ? this.children.length : index, 0, child);
    return child;
  }

  addEventListener(type: string, listener: (event: unknown) => void): void {
    this.listeners.set(type, [...(this.listeners.get(type) ?? []), listener]);
  }

  dispatch(type: string, event: Record<string, unknown>): void {
    for (const listener of this.listeners.get(type) ?? []) listener(event);
  }

  focus(): void {
    (globalThis.document as unknown as { activeElement: unknown }).activeElement = this;
  }

  contains(node: unknown): boolean {
    for (let current = node as FakeElement | null; current; current = current.parent) {
      if (current === this) return true;
    }
    return false;
  }

  closest(selector: string): FakeElement | null {
    for (let current: FakeElement | null = this; current; current = current.parent) {
      if (current.classList.contains(selector.slice(1))) return current;
    }
    return null;
  }

  querySelectorAll(selector: string): FakeElement[] {
    const matches: FakeElement[] = [];
    const visit = (node: FakeElement) => {
      for (const child of node.children) {
        if (child.classList.contains(selector.slice(1))) matches.push(child);
        visit(child);
      }
    };
    visit(this);
    return matches;
  }

  setAttribute(): void {}
  removeEventListener(): void {}
  select(): void {}
  scrollIntoView(): void {}
  getClientRects(): unknown[] {
    return [{}];
  }
}

const SESSION_BINDINGS: CompiledSessionBinding[] = [
  {
    sourcePath: 'stats.toggleKey',
    originalKey: 'Backquote',
    key: { code: 'Backquote', modifiers: [] },
    actionType: 'session-action',
    actionId: 'toggleStatsOverlay',
  },
  {
    sourcePath: 'shortcuts.toggleVisibleOverlayGlobal',
    originalKey: 'KeyO',
    key: { code: 'KeyO', modifiers: ['alt'] },
    actionType: 'session-action',
    actionId: 'toggleVisibleOverlay',
  },
];

function withFakeDom(
  run: (harness: ReturnType<typeof createHarness>) => Promise<void>,
  options: HarnessOptions = {},
) {
  return async () => {
    const globals = globalThis as Record<string, unknown>;
    const saved = ['window', 'document', 'HTMLElement', 'Element'].map((key) => [
      key,
      globals[key],
    ]);
    try {
      await run(createHarness(options));
    } finally {
      for (const [key, value] of saved) {
        Object.defineProperty(globalThis, key as string, {
          configurable: true,
          writable: true,
          value,
        });
      }
    }
  };
}

type HarnessOptions = { failActions?: boolean; bindings?: CompiledSessionBinding[] };

function createHarness(options: HarnessOptions) {
  const mpvCommands: (string | number)[][] = [];
  const sessionActions: string[] = [];
  /** IPC in send order; main destroys the modal window on `closed`, dropping anything after it. */
  const ipcLog: string[] = [];
  const define = (key: string, value: unknown) =>
    Object.defineProperty(globalThis, key, { configurable: true, writable: true, value });

  define('HTMLElement', FakeElement);
  define('Element', FakeElement);
  define('document', {
    activeElement: null,
    createElement: () => new FakeElement(),
    addEventListener: () => {},
    removeEventListener: () => {},
  });
  define('window', {
    electronAPI: {
      focusMainWindow: async () => {},
      setIgnoreMouseEvents: () => {},
      notifyOverlayModalClosed: () => ipcLog.push('closed'),
      getSessionBindings: async () => options.bindings ?? SESSION_BINDINGS,
      getSubtitleStyle: async () => ({}),
      getMarkWatchedKey: async () => null,
      getSubtitleSidebarSnapshot: async () => ({ config: { toggleKey: null } }),
      getRuntimeOptions: async () => [],
      getPlaybackPaused: async () => true,
      sendMpvCommand: (command: (string | number)[]) => {
        mpvCommands.push(command);
        ipcLog.push(`mpv:${command.join(' ')}`);
      },
      dispatchSessionAction: async (actionId: string) => {
        sessionActions.push(actionId);
        ipcLog.push(`action:${actionId}`);
        if (options.failActions) throw new Error('boom');
      },
    },
    focus: () => {},
    addEventListener: () => {},
    removeEventListener: () => {},
    setTimeout: (callback: () => void) => setTimeout(callback, 0),
    clearTimeout: (id: unknown) => clearTimeout(id as ReturnType<typeof setTimeout>),
  });

  const dom = {
    overlay: new FakeElement(),
    sessionHelpModal: new FakeElement(),
    sessionHelpFilter: new FakeElement(),
    sessionHelpContent: new FakeElement(),
    sessionHelpClose: new FakeElement(),
    sessionHelpShortcut: new FakeElement(),
    sessionHelpWarning: new FakeElement(),
    sessionHelpStatus: new FakeElement(),
  };
  const state = createRendererState();
  const modal = createSessionHelpModal(
    {
      state,
      dom,
      platform: {
        overlayLayer: 'modal',
        isModalLayer: true,
        isLinuxPlatform: true,
        isMacOSPlatform: false,
        isWindowsPlatform: false,
        shouldToggleMouseIgnore: false,
      },
    } as never,
    {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
    },
  );
  modal.wireDomEvents();

  async function open(commandsEnabled: boolean): Promise<void> {
    modal.openSessionHelpModal(
      { bindingKey: 'KeyH', fallbackUsed: false, fallbackUnavailable: false },
      { commandsEnabled },
    );
    await new Promise((resolve) => setTimeout(resolve, 0));
  }

  /** Keydown targets the focused element, as in the browser; returns whether default was prevented. */
  function pressEnter(target = (document as unknown as { activeElement: unknown }).activeElement) {
    let prevented = false;
    modal.handleSessionHelpKeydown({
      key: 'Enter',
      target,
      ctrlKey: false,
      metaKey: false,
      altKey: false,
      shiftKey: false,
      preventDefault: () => {
        prevented = true;
      },
    } as KeyboardEvent);
    return prevented;
  }

  const rows = () => dom.sessionHelpContent.querySelectorAll('.session-help-item');
  return { dom, state, open, pressEnter, rows, mpvCommands, sessionActions, ipcLog };
}

test('session help rows carry runnable commands except numeric and self-opening actions', () => {
  const sections = buildSessionHelpSections({
    sessionBindings: [
      {
        sourcePath: 'keybindings[0].key',
        originalKey: 'Space',
        key: { code: 'Space', modifiers: [] },
        actionType: 'mpv-command',
        command: ['cycle', 'pause'],
      },
      {
        sourcePath: 'shortcuts.copySubtitleMultiple',
        originalKey: 'Shift+KeyC',
        key: { code: 'KeyC', modifiers: ['shift'] },
        actionType: 'session-action',
        actionId: 'copySubtitleMultiple',
      },
      {
        sourcePath: 'shortcuts.openSessionHelp',
        originalKey: 'Slash',
        key: { code: 'Slash', modifiers: [] },
        actionType: 'session-action',
        actionId: 'openSessionHelp',
      },
    ],
    markWatchedKey: 'KeyW',
    subtitleStyle: {},
  });
  const rows = sections.flatMap((section) => section.rows);
  const commandFor = (action: string) => rows.find((row) => row.action === action)?.command;

  assert.deepEqual(commandFor('Toggle playback'), {
    actionType: 'mpv-command',
    command: ['cycle', 'pause'],
  });
  assert.deepEqual(commandFor('Mark video watched'), {
    actionType: 'session-action',
    actionId: 'markWatched',
  });
  assert.equal(commandFor('Copy subtitle (multi)'), undefined);
  assert.equal(commandFor('Open session help'), undefined);
  assert.equal(commandFor('Toggle primary subtitle bar visibility'), undefined);
});

test(
  'session help runs the selected command on Enter during playback and closes',
  withFakeDom(async (harness) => {
    await harness.open(true);
    assert.ok(harness.rows()[0]?.classList.contains('session-help-item-runnable'));

    harness.pressEnter();

    assert.deepEqual(harness.ipcLog, ['action:toggleStatsOverlay', 'closed']);
    assert.equal(harness.state.sessionHelpModalOpen, false);
  }),
);

test(
  'session help sends a paused subtitle seek before closing',
  withFakeDom(
    async (harness) => {
      await harness.open(true);
      // Playback rows live on another tab; a filter searches every tab.
      harness.dom.sessionHelpFilter.value = 'next subtitle';
      harness.dom.sessionHelpFilter.dispatch('input', {});

      harness.dom.sessionHelpContent.dispatch('dblclick', { target: harness.rows()[0] });
      await new Promise((resolve) => setTimeout(resolve, 0));

      assert.deepEqual(harness.ipcLog, ['mpv:sub-seek 1', 'mpv:set_property pause yes', 'closed']);
    },
    {
      bindings: [
        {
          sourcePath: 'keybindings[0].key',
          originalKey: 'Shift+KeyL',
          key: { code: 'KeyL', modifiers: ['shift'] },
          actionType: 'mpv-command',
          command: ['sub-seek', 1],
        },
      ],
    },
  ),
);

test(
  'session help ignores Enter when no video is playing',
  withFakeDom(async (harness) => {
    await harness.open(false);
    assert.equal(harness.rows()[0]?.classList.contains('session-help-item-runnable'), false);

    harness.pressEnter();

    assert.deepEqual(harness.sessionActions, []);
    assert.equal(harness.state.sessionHelpModalOpen, true);
  }),
);

test(
  'session help hover picks the Enter target and double-click runs the row',
  withFakeDom(async (harness) => {
    await harness.open(true);
    const [statsRow, overlayRow, fixedRow] = harness.rows();

    harness.dom.sessionHelpContent.dispatch('mousemove', { target: overlayRow });
    assert.equal(harness.state.sessionHelpSelectedIndex, 1);

    harness.dom.sessionHelpContent.dispatch('dblclick', { target: fixedRow });
    assert.equal(harness.state.sessionHelpModalOpen, true);

    harness.dom.sessionHelpContent.dispatch('dblclick', { target: statsRow?.children[0] });
    assert.deepEqual(harness.sessionActions, ['toggleStatsOverlay']);
    assert.equal(harness.state.sessionHelpModalOpen, false);
  }),
);

test(
  'session help reports a failed session action on the mpv OSD',
  withFakeDom(
    async (harness) => {
      const originalConsoleError = console.error;
      console.error = () => {};
      try {
        await harness.open(true);
        harness.pressEnter();
        await new Promise((resolve) => setTimeout(resolve, 0));
      } finally {
        console.error = originalConsoleError;
      }

      assert.deepEqual(harness.mpvCommands, [['show-text', 'Command failed to run', '3000']]);
    },
    { failActions: true },
  ),
);

test(
  'session help does not run a command from Enter pressed in the filter',
  withFakeDom(async (harness) => {
    await harness.open(true);
    // The filter's own listener already moved focus to a row before this bubbles up.
    harness.rows()[0]?.focus();

    harness.pressEnter(harness.dom.sessionHelpFilter);

    assert.deepEqual(harness.sessionActions, []);
    assert.equal(harness.state.sessionHelpModalOpen, true);
  }),
);

test(
  'session help leaves Enter on the close button to its native activation',
  withFakeDom(async (harness) => {
    await harness.open(true);

    const prevented = harness.pressEnter(harness.dom.sessionHelpClose);

    assert.equal(prevented, false);
    assert.deepEqual(harness.sessionActions, []);
    assert.equal(harness.state.sessionHelpModalOpen, true);
  }),
);

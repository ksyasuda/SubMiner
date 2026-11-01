import assert from 'node:assert/strict';
import test from 'node:test';
import type { GlobalShortcut } from 'electron';
import { DEFAULT_CONFIG } from '../../config/definitions';
import { createShortcutsAssembly, type ShortcutsAssemblyDeps } from './shortcuts';

// Every callback throws, so a test fails if the assembly calls a dep it should only call later.
function makeDeps(overrides: Partial<ShortcutsAssemblyDeps> = {}): ShortcutsAssemblyDeps {
  const lazy = (name: string) => (): never => {
    throw new Error(`${name} called unexpectedly`);
  };
  return {
    appState: {
      shortcutsRegistered: false,
      overlayRuntimeInitialized: false,
      windowTracker: null,
      keybindings: [],
      mpvClient: null,
      sessionBindings: [],
      sessionBindingsInitialized: false,
    },
    getConfig: lazy('getConfig'),
    logger: { warn: lazy('warn') },
    configDir: '/tmp/subminer-shortcuts-assembly-test',
    isDev: false,
    overlayManager: {
      getMainWindow: lazy('getMainWindow'),
      getVisibleOverlayVisible: lazy('getVisibleOverlayVisible'),
      broadcastToOverlayWindows: lazy('broadcastToOverlayWindows'),
    },
    globalShortcut: new Proxy({}, { get: (_target, name: string) => lazy(name) }) as GlobalShortcut,
    getOverlayNotifications: lazy('getOverlayNotifications'),
    getOverlayModals: lazy('getOverlayModals'),
    getMining: lazy('getMining'),
    triggerSubsyncFromConfig: lazy('triggerSubsyncFromConfig'),
    cycleSecondarySubMode: lazy('cycleSecondarySubMode'),
    toggleVisibleOverlay: lazy('toggleVisibleOverlay'),
    openDictionarySettings: lazy('openDictionarySettings'),
    ...overrides,
  };
}

// main.ts builds the assembly before most siblings exist, so construction must stay lazy.
test('createShortcutsAssembly does not call its deps while constructing', () => {
  const shortcuts = createShortcutsAssembly(makeDeps());
  assert.equal(typeof shortcuts.registerGlobalShortcuts, 'function');
  assert.equal(typeof shortcuts.syncOverlayShortcutsForModal, 'function');
});

test('opening a modal releases the registered overlay shortcuts', () => {
  const deps = makeDeps({ getConfig: () => DEFAULT_CONFIG });
  deps.appState.shortcutsRegistered = true;
  const shortcuts = createShortcutsAssembly(deps);

  shortcuts.syncOverlayShortcutsForModal(true);

  assert.equal(deps.appState.shortcutsRegistered, false);
});

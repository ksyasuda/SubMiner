import assert from 'node:assert/strict';
import test from 'node:test';
import type { Keybinding } from '../../types';
import type { ConfiguredShortcuts } from '../utils/shortcut-config';
import { SPECIAL_COMMANDS } from '../../config/definitions';
import { compileSessionBindings } from './session-bindings';

function createShortcuts(overrides: Partial<ConfiguredShortcuts> = {}): ConfiguredShortcuts {
  return {
    toggleVisibleOverlayGlobal: null,
    copySubtitle: null,
    copySubtitleMultiple: null,
    updateLastCardFromClipboard: null,
    triggerFieldGrouping: null,
    triggerSubsync: null,
    mineSentence: null,
    mineSentenceMultiple: null,
    multiCopyTimeoutMs: 2500,
    toggleSecondarySub: null,
    markAudioCard: null,
    openRuntimeOptions: null,
    openJimaku: null,
    ...overrides,
  };
}

function createKeybinding(key: string, command: Keybinding['command']): Keybinding {
  return { key, command };
}

test('compileSessionBindings merges shortcuts and keybindings into one canonical list', () => {
  const result = compileSessionBindings({
    shortcuts: createShortcuts({
      toggleVisibleOverlayGlobal: 'Alt+Shift+O',
      openJimaku: 'Ctrl+Shift+J',
    }),
    keybindings: [
      createKeybinding('KeyF', ['cycle', 'fullscreen']),
      createKeybinding('Ctrl+Shift+Y', [SPECIAL_COMMANDS.YOUTUBE_PICKER_OPEN]),
    ],
    platform: 'linux',
  });

  assert.equal(result.warnings.length, 0);
  assert.deepEqual(
    result.bindings.map((binding) => ({
      actionType: binding.actionType,
      sourcePath: binding.sourcePath,
      code: binding.key.code,
      modifiers: binding.key.modifiers,
      target:
        binding.actionType === 'session-action'
          ? binding.actionId
          : binding.command.join(' '),
    })),
    [
      {
        actionType: 'mpv-command',
        sourcePath: 'keybindings[0].key',
        code: 'KeyF',
        modifiers: [],
        target: 'cycle fullscreen',
      },
      {
        actionType: 'session-action',
        sourcePath: 'keybindings[1].key',
        code: 'KeyY',
        modifiers: ['ctrl', 'shift'],
        target: 'openYoutubePicker',
      },
      {
        actionType: 'session-action',
        sourcePath: 'shortcuts.openJimaku',
        code: 'KeyJ',
        modifiers: ['ctrl', 'shift'],
        target: 'openJimaku',
      },
      {
        actionType: 'session-action',
        sourcePath: 'shortcuts.toggleVisibleOverlayGlobal',
        code: 'KeyO',
        modifiers: ['alt', 'shift'],
        target: 'toggleVisibleOverlay',
      },
    ],
  );
});

test('compileSessionBindings resolves CommandOrControl per platform', () => {
  const input = {
    shortcuts: createShortcuts({
      toggleVisibleOverlayGlobal: 'CommandOrControl+Shift+O',
    }),
    keybindings: [],
  };

  const windows = compileSessionBindings({ ...input, platform: 'win32' });
  const mac = compileSessionBindings({ ...input, platform: 'darwin' });

  assert.deepEqual(windows.bindings[0]?.key.modifiers, ['ctrl', 'shift']);
  assert.deepEqual(mac.bindings[0]?.key.modifiers, ['shift', 'meta']);
});

test('compileSessionBindings drops conflicting bindings that canonicalize to the same key', () => {
  const result = compileSessionBindings({
    shortcuts: createShortcuts({
      openJimaku: 'Ctrl+Shift+J',
    }),
    keybindings: [createKeybinding('Ctrl+Shift+J', [SPECIAL_COMMANDS.PLAYLIST_BROWSER_OPEN])],
    platform: 'linux',
  });

  assert.deepEqual(result.bindings, []);
  assert.equal(result.warnings.length, 1);
  assert.equal(result.warnings[0]?.kind, 'conflict');
  assert.deepEqual(result.warnings[0]?.conflictingPaths, [
    'shortcuts.openJimaku',
    'keybindings[0].key',
  ]);
});

test('compileSessionBindings omits disabled bindings', () => {
  const result = compileSessionBindings({
    shortcuts: createShortcuts({
      openJimaku: null,
      toggleVisibleOverlayGlobal: 'Alt+Shift+O',
    }),
    keybindings: [createKeybinding('Ctrl+Shift+J', null)],
    platform: 'linux',
  });

  assert.equal(result.warnings.length, 0);
  assert.deepEqual(result.bindings.map((binding) => binding.sourcePath), [
    'shortcuts.toggleVisibleOverlayGlobal',
  ]);
});

test('compileSessionBindings warns on unsupported shortcut and keybinding syntax', () => {
  const result = compileSessionBindings({
    shortcuts: createShortcuts({
      openJimaku: 'Hyper+J',
    }),
    keybindings: [createKeybinding('Ctrl+ß', ['cycle', 'fullscreen'])],
    platform: 'linux',
  });

  assert.deepEqual(result.bindings, []);
  assert.deepEqual(
    result.warnings.map((warning) => `${warning.kind}:${warning.path}`),
    ['unsupported:shortcuts.openJimaku', 'unsupported:keybindings[0].key'],
  );
});

test('compileSessionBindings warns on deprecated toggleVisibleOverlayGlobal config', () => {
  const result = compileSessionBindings({
    shortcuts: createShortcuts(),
    keybindings: [],
    platform: 'linux',
    rawConfig: {
      shortcuts: {
        toggleVisibleOverlayGlobal: 'Alt+Shift+O',
      },
    } as never,
  });

  assert.equal(result.bindings.length, 0);
  assert.deepEqual(result.warnings, [
    {
      kind: 'deprecated-config',
      path: 'shortcuts.toggleVisibleOverlayGlobal',
      value: 'Alt+Shift+O',
      message: 'Rename shortcuts.toggleVisibleOverlayGlobal to shortcuts.toggleVisibleOverlay.',
    },
  ]);
});

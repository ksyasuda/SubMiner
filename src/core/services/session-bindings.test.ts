import assert from 'node:assert/strict';
import test from 'node:test';
import type { Keybinding } from '../../types';
import type { ConfiguredShortcuts } from '../utils/shortcut-config';
import { DEFAULT_CONFIG, DEFAULT_KEYBINDINGS, SPECIAL_COMMANDS } from '../../config/definitions';
import { resolveConfiguredShortcuts } from '../utils/shortcut-config';
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
    openCharacterDictionary: null,
    openRuntimeOptions: null,
    openJimaku: null,
    openSessionHelp: null,
    openControllerSelect: null,
    openControllerDebug: null,
    toggleSubtitleSidebar: null,
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
      openControllerSelect: 'Alt+C',
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
        binding.actionType === 'session-action' ? binding.actionId : binding.command.join(' '),
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
        sourcePath: 'shortcuts.openControllerSelect',
        code: 'KeyC',
        modifiers: ['alt'],
        target: 'openControllerSelect',
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

test('compileSessionBindings resolves CommandOrControl in DOM key strings per platform', () => {
  const input = {
    shortcuts: createShortcuts(),
    keybindings: [createKeybinding('CommandOrControl+Shift+J', ['cycle', 'fullscreen'])],
    statsToggleKey: 'CommandOrControl+Backquote',
  };

  const windows = compileSessionBindings({ ...input, platform: 'win32' });
  const mac = compileSessionBindings({ ...input, platform: 'darwin' });

  assert.deepEqual(
    windows.bindings
      .map((binding) => ({
        sourcePath: binding.sourcePath,
        modifiers: binding.key.modifiers,
      }))
      .sort((left, right) => left.sourcePath.localeCompare(right.sourcePath)),
    [
      {
        sourcePath: 'keybindings[0].key',
        modifiers: ['ctrl', 'shift'],
      },
      {
        sourcePath: 'stats.toggleKey',
        modifiers: ['ctrl'],
      },
    ],
  );

  assert.deepEqual(
    mac.bindings
      .map((binding) => ({
        sourcePath: binding.sourcePath,
        modifiers: binding.key.modifiers,
      }))
      .sort((left, right) => left.sourcePath.localeCompare(right.sourcePath)),
    [
      {
        sourcePath: 'keybindings[0].key',
        modifiers: ['shift', 'meta'],
      },
      {
        sourcePath: 'stats.toggleKey',
        modifiers: ['meta'],
      },
    ],
  );
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

test('compileSessionBindings keeps default replay and next subtitle session actions on Linux', () => {
  const result = compileSessionBindings({
    shortcuts: resolveConfiguredShortcuts(DEFAULT_CONFIG, DEFAULT_CONFIG),
    keybindings: DEFAULT_KEYBINDINGS,
    statsToggleKey: DEFAULT_CONFIG.stats.toggleKey,
    platform: 'linux',
    rawConfig: DEFAULT_CONFIG,
  });

  assert.deepEqual(
    result.warnings.filter((warning) => warning.kind === 'conflict'),
    [],
  );
  const bySignature = new Map(
    result.bindings.map((binding) => [
      `${binding.key.modifiers.join('+')}+${binding.key.code}`,
      binding,
    ]),
  );

  const replay = bySignature.get('ctrl+shift+KeyH');
  assert.equal(replay?.actionType, 'session-action');
  assert.equal(replay?.actionId, 'replayCurrentSubtitle');

  const next = bySignature.get('ctrl+shift+KeyL');
  assert.equal(next?.actionType, 'session-action');
  assert.equal(next?.actionId, 'playNextSubtitle');
});

test('compileSessionBindings wires every default keybinding to an overlay or mpv action', () => {
  const expectedSpecialActions: Record<string, string> = {
    [SPECIAL_COMMANDS.SHIFT_SUB_DELAY_TO_PREVIOUS_SUBTITLE_START]: 'shiftSubDelayPrevLine',
    [SPECIAL_COMMANDS.SHIFT_SUB_DELAY_TO_NEXT_SUBTITLE_START]: 'shiftSubDelayNextLine',
    [SPECIAL_COMMANDS.YOUTUBE_PICKER_OPEN]: 'openYoutubePicker',
    [SPECIAL_COMMANDS.PLAYLIST_BROWSER_OPEN]: 'openPlaylistBrowser',
    [SPECIAL_COMMANDS.REPLAY_SUBTITLE]: 'replayCurrentSubtitle',
    [SPECIAL_COMMANDS.PLAY_NEXT_SUBTITLE]: 'playNextSubtitle',
  };
  const result = compileSessionBindings({
    shortcuts: createShortcuts(),
    keybindings: DEFAULT_KEYBINDINGS,
    platform: 'linux',
  });

  assert.deepEqual(result.warnings, []);
  const byOriginalKey = new Map(result.bindings.map((binding) => [binding.originalKey, binding]));
  assert.equal(byOriginalKey.size, DEFAULT_KEYBINDINGS.length);

  for (const defaultBinding of DEFAULT_KEYBINDINGS) {
    const compiled = byOriginalKey.get(defaultBinding.key);
    assert.ok(compiled, `${defaultBinding.key} should compile`);

    const specialAction = expectedSpecialActions[String(defaultBinding.command?.[0])];
    if (specialAction) {
      assert.equal(compiled.actionType, 'session-action');
      assert.equal(compiled.actionId, specialAction);
      continue;
    }

    assert.equal(compiled.actionType, 'mpv-command');
    assert.deepEqual(compiled.command, defaultBinding.command);
  }
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
  assert.deepEqual(
    result.bindings.map((binding) => binding.sourcePath),
    ['shortcuts.toggleVisibleOverlayGlobal'],
  );
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

test('compileSessionBindings rejects malformed command arrays', () => {
  const result = compileSessionBindings({
    shortcuts: createShortcuts(),
    keybindings: [
      createKeybinding('Ctrl+J', ['show-text', 3000]),
      createKeybinding('Ctrl+K', ['show-text', { bad: true } as never] as never),
    ],
    platform: 'linux',
  });

  assert.deepEqual(
    result.bindings.map((binding) => binding.sourcePath),
    ['keybindings[0].key'],
  );
  assert.equal(result.bindings[0]?.actionType, 'mpv-command');
  assert.deepEqual(result.bindings[0]?.command, ['show-text', 3000]);
  assert.deepEqual(
    result.warnings.map((warning) => `${warning.kind}:${warning.path}`),
    ['unsupported:keybindings[1].command'],
  );
});

test('compileSessionBindings rejects non-string command heads and extra args on special commands', () => {
  const result = compileSessionBindings({
    shortcuts: createShortcuts(),
    keybindings: [
      createKeybinding('Ctrl+J', [42] as never),
      createKeybinding('Ctrl+K', [SPECIAL_COMMANDS.JIMAKU_OPEN, 'extra'] as never),
    ],
    platform: 'linux',
  });

  assert.deepEqual(result.bindings, []);
  assert.deepEqual(
    result.warnings.map((warning) => `${warning.kind}:${warning.path}`),
    ['unsupported:keybindings[0].command', 'unsupported:keybindings[1].command'],
  );
});

test('compileSessionBindings points unsupported command warnings at the command field', () => {
  const result = compileSessionBindings({
    shortcuts: createShortcuts(),
    keybindings: [createKeybinding('Ctrl+K', [SPECIAL_COMMANDS.JIMAKU_OPEN, 'extra'] as never)],
    platform: 'linux',
  });

  assert.deepEqual(result.bindings, []);
  assert.deepEqual(
    result.warnings.map((warning) => `${warning.kind}:${warning.path}`),
    ['unsupported:keybindings[0].command'],
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

test('compileSessionBindings includes stats toggle in the shared session binding artifact', () => {
  const result = compileSessionBindings({
    shortcuts: createShortcuts(),
    keybindings: [],
    statsToggleKey: 'Backquote',
    platform: 'win32',
  });

  assert.equal(result.warnings.length, 0);
  assert.deepEqual(result.bindings, [
    {
      sourcePath: 'stats.toggleKey',
      originalKey: 'Backquote',
      key: {
        code: 'Backquote',
        modifiers: [],
      },
      actionType: 'session-action',
      actionId: 'toggleStatsOverlay',
    },
  ]);
});

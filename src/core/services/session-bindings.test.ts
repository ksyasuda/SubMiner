import assert from 'node:assert/strict';
import test from 'node:test';
import type { Keybinding } from '../../types';
import type { ConfiguredShortcuts } from '../utils/shortcut-config';
import { DEFAULT_CONFIG, DEFAULT_KEYBINDINGS, SPECIAL_COMMANDS } from '../../config/definitions';
import { resolveConfiguredShortcuts } from '../utils/shortcut-config';
import { buildPluginSessionBindingsArtifact, compileSessionBindings } from './session-bindings';

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
    openCharacterDictionaryManager: null,
    openRuntimeOptions: null,
    openJimaku: null,
    openSessionHelp: null,
    openControllerSelect: null,
    openControllerDebug: null,
    toggleSubtitleSidebar: null,
    toggleNotificationHistory: null,
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

test('compileSessionBindings supports mpv mouse button keybindings', () => {
  const result = compileSessionBindings({
    shortcuts: createShortcuts(),
    keybindings: [
      createKeybinding('MBTN_BACK', ['sub-seek', -1]),
      createKeybinding('Shift+MBTN_FORWARD', ['sub-seek', 1]),
    ],
    platform: 'win32',
  });

  assert.deepEqual(result.warnings, []);
  assert.deepEqual(
    result.bindings.map((binding) => ({
      code: binding.key.code,
      modifiers: binding.key.modifiers,
      command: binding.actionType === 'mpv-command' ? binding.command : null,
    })),
    [
      { code: 'MBTN_BACK', modifiers: [], command: ['sub-seek', -1] },
      { code: 'MBTN_FORWARD', modifiers: ['shift'], command: ['sub-seek', 1] },
    ],
  );
});

test('compileSessionBindings keeps mouse buttons scoped to keybindings', () => {
  const result = compileSessionBindings({
    shortcuts: createShortcuts({
      openJimaku: 'MBTN_BACK',
    }),
    keybindings: [createKeybinding('MBTN_BACK', ['sub-seek', -1])],
    platform: 'win32',
  });

  assert.deepEqual(
    result.bindings.map((binding) => binding.sourcePath),
    ['keybindings[0].key'],
  );
  assert.deepEqual(
    result.warnings.map((warning) => `${warning.kind}:${warning.path}`),
    ['unsupported:shortcuts.openJimaku'],
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

test('compileSessionBindings keeps only the character dictionary manager bound by default', () => {
  const result = compileSessionBindings({
    shortcuts: resolveConfiguredShortcuts(DEFAULT_CONFIG, DEFAULT_CONFIG),
    keybindings: DEFAULT_KEYBINDINGS,
    statsToggleKey: DEFAULT_CONFIG.stats.toggleKey,
    platform: 'linux',
    rawConfig: DEFAULT_CONFIG,
  });

  const characterDictionaryBindings = result.bindings.flatMap((binding) => {
    if (binding.actionType !== 'session-action') return [];
    if (binding.actionId !== 'openCharacterDictionaryManager') {
      return [];
    }
    return [
      {
        sourcePath: binding.sourcePath,
        originalKey: binding.originalKey,
        actionId: binding.actionId,
      },
    ];
  });

  assert.deepEqual(characterDictionaryBindings, [
    {
      sourcePath: 'shortcuts.openCharacterDictionaryManager',
      originalKey: 'CommandOrControl+D',
      actionId: 'openCharacterDictionaryManager',
    },
  ]);
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

test('compileSessionBindings includes mark-watched in the shared session binding artifact', () => {
  const result = compileSessionBindings({
    shortcuts: createShortcuts(),
    keybindings: [],
    statsMarkWatchedKey: 'Ctrl+Shift+KeyW',
    platform: 'darwin',
  });

  assert.equal(result.warnings.length, 0);
  assert.deepEqual(result.bindings, [
    {
      sourcePath: 'stats.markWatchedKey',
      originalKey: 'Ctrl+Shift+KeyW',
      key: {
        code: 'KeyW',
        modifiers: ['ctrl', 'shift'],
      },
      actionType: 'session-action',
      actionId: 'markWatched',
    },
  ]);
});

test('compileSessionBindings wires every configured shortcut key into the shared artifact', () => {
  const shortcutKeys: Array<keyof Omit<ConfiguredShortcuts, 'multiCopyTimeoutMs'>> = [
    'toggleVisibleOverlayGlobal',
    'copySubtitle',
    'copySubtitleMultiple',
    'updateLastCardFromClipboard',
    'triggerFieldGrouping',
    'triggerSubsync',
    'mineSentence',
    'mineSentenceMultiple',
    'toggleSecondarySub',
    'markAudioCard',
    'openCharacterDictionaryManager',
    'openRuntimeOptions',
    'openJimaku',
    'openSessionHelp',
    'openControllerSelect',
    'openControllerDebug',
    'toggleSubtitleSidebar',
  ];
  const shortcuts = createShortcuts();
  shortcutKeys.forEach((key, index) => {
    shortcuts[key] = `Ctrl+Alt+F${index + 1}`;
  });

  const result = compileSessionBindings({
    shortcuts,
    keybindings: [],
    platform: 'linux',
  });

  assert.deepEqual(result.warnings, []);
  assert.deepEqual(
    result.bindings.map((binding) => binding.sourcePath).sort(),
    shortcutKeys.map((key) => `shortcuts.${key}`).sort(),
  );
});

test('buildPluginSessionBindingsArtifact emits CLI args for plugin-bound session actions', () => {
  const result = compileSessionBindings({
    shortcuts: createShortcuts({
      openCharacterDictionaryManager: 'Ctrl+D',
    }),
    keybindings: [
      createKeybinding('Ctrl+Alt+KeyR', [
        `${SPECIAL_COMMANDS.RUNTIME_OPTION_CYCLE_PREFIX}anki.autoUpdateNewCards:prev`,
      ]),
    ],
    platform: 'linux',
  });

  const artifact = buildPluginSessionBindingsArtifact({
    bindings: result.bindings,
    warnings: result.warnings,
    numericSelectionTimeoutMs: 2500,
    now: new Date('2026-05-26T00:00:00.000Z'),
  });
  const byActionId = new Map(
    artifact.bindings.flatMap((binding) =>
      binding.actionType === 'session-action' ? [[binding.actionId, binding]] : [],
    ),
  );
  const compiledManagerBinding = result.bindings.find(
    (binding) =>
      binding.actionType === 'session-action' &&
      binding.actionId === 'openCharacterDictionaryManager',
  );

  assert.equal(compiledManagerBinding && 'cliArgs' in compiledManagerBinding, false);
  const managerCliArgs = byActionId.get('openCharacterDictionaryManager')?.cliArgs;
  const cycleCliArgs = byActionId.get('cycleRuntimeOption')?.cliArgs;

  assert.equal(managerCliArgs?.[0], '--session-action');
  assert.deepEqual(JSON.parse(managerCliArgs?.[1] ?? ''), {
    actionId: 'openCharacterDictionaryManager',
  });
  assert.equal(cycleCliArgs?.[0], '--session-action');
  assert.deepEqual(JSON.parse(cycleCliArgs?.[1] ?? ''), {
    actionId: 'cycleRuntimeOption',
    payload: {
      runtimeOptionId: 'anki.autoUpdateNewCards',
      direction: -1,
    },
  });
});

test('buildPluginSessionBindingsArtifact preserves plugin selector CLI for no-count multi-line actions', () => {
  const result = compileSessionBindings({
    shortcuts: createShortcuts({
      copySubtitleMultiple: 'Ctrl+Shift+C',
      mineSentenceMultiple: 'Ctrl+Shift+S',
    }),
    keybindings: [],
    platform: 'linux',
  });

  const artifact = buildPluginSessionBindingsArtifact({
    bindings: result.bindings,
    warnings: result.warnings,
    numericSelectionTimeoutMs: 2500,
  });
  const byActionId = new Map(
    artifact.bindings.flatMap((binding) =>
      binding.actionType === 'session-action' ? [[binding.actionId, binding]] : [],
    ),
  );

  assert.equal(byActionId.get('copySubtitleMultiple')?.cliArgs, undefined);
  assert.equal(byActionId.get('mineSentenceMultiple')?.cliArgs, undefined);
});

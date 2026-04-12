import test from 'node:test';
import assert from 'node:assert/strict';
import { ConfiguredShortcuts } from '../utils/shortcut-config';
import {
  createOverlayShortcutRuntimeHandlers,
  OverlayShortcutRuntimeDeps,
  runOverlayShortcutLocalFallback,
} from './overlay-shortcut-handler';
import {
  registerOverlayShortcutsRuntime,
  shouldActivateOverlayShortcuts,
  unregisterOverlayShortcutsRuntime,
} from './overlay-shortcut';

function makeShortcuts(overrides: Partial<ConfiguredShortcuts> = {}): ConfiguredShortcuts {
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
    openSessionHelp: null,
    openControllerSelect: null,
    openControllerDebug: null,
    toggleSubtitleSidebar: null,
    ...overrides,
  };
}

function createDeps(overrides: Partial<OverlayShortcutRuntimeDeps> = {}) {
  const calls: string[] = [];
  const osd: string[] = [];
  const deps: OverlayShortcutRuntimeDeps = {
    showMpvOsd: (text) => {
      osd.push(text);
    },
    openRuntimeOptions: () => {
      calls.push('openRuntimeOptions');
    },
    openJimaku: () => {
      calls.push('openJimaku');
    },
    markAudioCard: async () => {
      calls.push('markAudioCard');
    },
    copySubtitleMultiple: (timeoutMs) => {
      calls.push(`copySubtitleMultiple:${timeoutMs}`);
    },
    copySubtitle: () => {
      calls.push('copySubtitle');
    },
    toggleSecondarySub: () => {
      calls.push('toggleSecondarySub');
    },
    updateLastCardFromClipboard: async () => {
      calls.push('updateLastCardFromClipboard');
    },
    triggerFieldGrouping: async () => {
      calls.push('triggerFieldGrouping');
    },
    triggerSubsync: async () => {
      calls.push('triggerSubsync');
    },
    mineSentence: async () => {
      calls.push('mineSentence');
    },
    mineSentenceMultiple: (timeoutMs) => {
      calls.push(`mineSentenceMultiple:${timeoutMs}`);
    },
    ...overrides,
  };

  return { deps, calls, osd };
}

test('createOverlayShortcutRuntimeHandlers dispatches sync and async handlers', async () => {
  const { deps, calls } = createDeps();
  const { overlayHandlers, fallbackHandlers } = createOverlayShortcutRuntimeHandlers(deps);

  overlayHandlers.copySubtitle();
  overlayHandlers.copySubtitleMultiple(1111);
  overlayHandlers.toggleSecondarySub();
  overlayHandlers.openRuntimeOptions();
  overlayHandlers.openJimaku();
  overlayHandlers.mineSentenceMultiple(2222);
  overlayHandlers.updateLastCardFromClipboard();
  fallbackHandlers.mineSentence();
  await new Promise((resolve) => setImmediate(resolve));

  assert.deepEqual(calls, [
    'copySubtitle',
    'copySubtitleMultiple:1111',
    'toggleSecondarySub',
    'openRuntimeOptions',
    'openJimaku',
    'mineSentenceMultiple:2222',
    'updateLastCardFromClipboard',
    'mineSentence',
  ]);
});

test('createOverlayShortcutRuntimeHandlers reports async failures via OSD', async () => {
  const logs: unknown[][] = [];
  const originalError = console.error;
  console.error = (...args: unknown[]) => {
    logs.push(args);
  };

  try {
    const { deps, osd } = createDeps({
      markAudioCard: async () => {
        throw new Error('audio boom');
      },
    });
    const { overlayHandlers } = createOverlayShortcutRuntimeHandlers(deps);

    overlayHandlers.markAudioCard();
    await new Promise((resolve) => setImmediate(resolve));

    assert.equal(logs.length, 1);
    assert.equal(typeof logs[0]?.[0], 'string');
    assert.ok(String(logs[0]?.[0]).includes('markLastCardAsAudioCard failed:'));
    assert.ok(String(logs[0]?.[0]).includes('audio boom'));
    assert.ok(osd.some((entry) => entry.includes('Audio card failed: audio boom')));
  } finally {
    console.error = originalError;
  }
});

test('runOverlayShortcutLocalFallback dispatches matching actions with timeout', () => {
  const handled: string[] = [];
  const matched: Array<{ accelerator: string; allowWhenRegistered: boolean }> = [];
  const shortcuts = makeShortcuts({
    copySubtitleMultiple: 'Ctrl+M',
    multiCopyTimeoutMs: 4321,
  });

  const result = runOverlayShortcutLocalFallback(
    {} as Electron.Input,
    shortcuts,
    (_input, accelerator, allowWhenRegistered) => {
      matched.push({
        accelerator,
        allowWhenRegistered: allowWhenRegistered === true,
      });
      return accelerator === 'Ctrl+M';
    },
    {
      openRuntimeOptions: () => handled.push('openRuntimeOptions'),
      openJimaku: () => handled.push('openJimaku'),
      markAudioCard: () => handled.push('markAudioCard'),
      copySubtitleMultiple: (timeoutMs) => handled.push(`copySubtitleMultiple:${timeoutMs}`),
      copySubtitle: () => handled.push('copySubtitle'),
      toggleSecondarySub: () => handled.push('toggleSecondarySub'),
      updateLastCardFromClipboard: () => handled.push('updateLastCardFromClipboard'),
      triggerFieldGrouping: () => handled.push('triggerFieldGrouping'),
      triggerSubsync: () => handled.push('triggerSubsync'),
      mineSentence: () => handled.push('mineSentence'),
      mineSentenceMultiple: (timeoutMs) => handled.push(`mineSentenceMultiple:${timeoutMs}`),
    },
  );

  assert.equal(result, true);
  assert.deepEqual(handled, ['copySubtitleMultiple:4321']);
  assert.deepEqual(matched, [{ accelerator: 'Ctrl+M', allowWhenRegistered: false }]);
});

test('runOverlayShortcutLocalFallback passes allowWhenRegistered for secondary-sub toggle', () => {
  const matched: Array<{ accelerator: string; allowWhenRegistered: boolean }> = [];
  const shortcuts = makeShortcuts({
    toggleSecondarySub: 'Ctrl+2',
  });

  const result = runOverlayShortcutLocalFallback(
    {} as Electron.Input,
    shortcuts,
    (_input, accelerator, allowWhenRegistered) => {
      matched.push({
        accelerator,
        allowWhenRegistered: allowWhenRegistered === true,
      });
      return accelerator === 'Ctrl+2';
    },
    {
      openRuntimeOptions: () => {},
      openJimaku: () => {},
      markAudioCard: () => {},
      copySubtitleMultiple: () => {},
      copySubtitle: () => {},
      toggleSecondarySub: () => {},
      updateLastCardFromClipboard: () => {},
      triggerFieldGrouping: () => {},
      triggerSubsync: () => {},
      mineSentence: () => {},
      mineSentenceMultiple: () => {},
    },
  );

  assert.equal(result, true);
  assert.deepEqual(matched, [{ accelerator: 'Ctrl+2', allowWhenRegistered: true }]);
});

test('runOverlayShortcutLocalFallback allows registered-global jimaku shortcut', () => {
  const matched: Array<{ accelerator: string; allowWhenRegistered: boolean }> = [];
  const shortcuts = makeShortcuts({
    openJimaku: 'Ctrl+J',
  });

  const result = runOverlayShortcutLocalFallback(
    {} as Electron.Input,
    shortcuts,
    (_input, accelerator, allowWhenRegistered) => {
      matched.push({
        accelerator,
        allowWhenRegistered: allowWhenRegistered === true,
      });
      return accelerator === 'Ctrl+J';
    },
    {
      openRuntimeOptions: () => {},
      openJimaku: () => {},
      markAudioCard: () => {},
      copySubtitleMultiple: () => {},
      copySubtitle: () => {},
      toggleSecondarySub: () => {},
      updateLastCardFromClipboard: () => {},
      triggerFieldGrouping: () => {},
      triggerSubsync: () => {},
      mineSentence: () => {},
      mineSentenceMultiple: () => {},
    },
  );

  assert.equal(result, true);
  assert.deepEqual(matched, [{ accelerator: 'Ctrl+J', allowWhenRegistered: true }]);
});

test('runOverlayShortcutLocalFallback returns false when no action matches', () => {
  const shortcuts = makeShortcuts({
    copySubtitle: 'Ctrl+C',
  });
  let called = false;

  const result = runOverlayShortcutLocalFallback({} as Electron.Input, shortcuts, () => false, {
    openRuntimeOptions: () => {
      called = true;
    },
    openJimaku: () => {
      called = true;
    },
    markAudioCard: () => {
      called = true;
    },
    copySubtitleMultiple: () => {
      called = true;
    },
    copySubtitle: () => {
      called = true;
    },
    toggleSecondarySub: () => {
      called = true;
    },
    updateLastCardFromClipboard: () => {
      called = true;
    },
    triggerFieldGrouping: () => {
      called = true;
    },
    triggerSubsync: () => {
      called = true;
    },
    mineSentence: () => {
      called = true;
    },
    mineSentenceMultiple: () => {
      called = true;
    },
  });

  assert.equal(result, false);
  assert.equal(called, false);
});

test('shouldActivateOverlayShortcuts disables macOS overlay shortcuts when tracked mpv is unfocused', () => {
  assert.equal(
    shouldActivateOverlayShortcuts({
      overlayRuntimeInitialized: true,
      isMacOSPlatform: true,
      trackedMpvWindowFocused: false,
    }),
    false,
  );
});

test('shouldActivateOverlayShortcuts keeps macOS overlay shortcuts active when tracked mpv is focused', () => {
  assert.equal(
    shouldActivateOverlayShortcuts({
      overlayRuntimeInitialized: true,
      isMacOSPlatform: true,
      trackedMpvWindowFocused: true,
    }),
    true,
  );
});

test('shouldActivateOverlayShortcuts preserves non-macOS behavior', () => {
  assert.equal(
    shouldActivateOverlayShortcuts({
      overlayRuntimeInitialized: true,
      isMacOSPlatform: false,
      trackedMpvWindowFocused: false,
    }),
    true,
  );
});

test('registerOverlayShortcutsRuntime reports active shortcuts when configured', () => {
  const deps = {
    getConfiguredShortcuts: () => makeShortcuts({ openJimaku: 'Ctrl+J' }),
    getOverlayHandlers: () => ({
      copySubtitle: () => {},
      copySubtitleMultiple: () => {},
      updateLastCardFromClipboard: () => {},
      triggerFieldGrouping: () => {},
      triggerSubsync: () => {},
      mineSentence: () => {},
      mineSentenceMultiple: () => {},
      toggleSecondarySub: () => {},
      markAudioCard: () => {},
      openRuntimeOptions: () => {},
      openJimaku: () => {},
    }),
    cancelPendingMultiCopy: () => {},
    cancelPendingMineSentenceMultiple: () => {},
  };

  const result = registerOverlayShortcutsRuntime(deps);
  assert.equal(result, true);
  assert.equal(unregisterOverlayShortcutsRuntime(result, deps), false);
});

test('unregisterOverlayShortcutsRuntime clears pending shortcut work when active', () => {
  const calls: string[] = [];
  const deps = {
    getConfiguredShortcuts: () => makeShortcuts({ openJimaku: 'Ctrl+J' }),
    getOverlayHandlers: () => ({
      copySubtitle: () => {},
      copySubtitleMultiple: () => {},
      updateLastCardFromClipboard: () => {},
      triggerFieldGrouping: () => {},
      triggerSubsync: () => {},
      mineSentence: () => {},
      mineSentenceMultiple: () => {},
      toggleSecondarySub: () => {},
      markAudioCard: () => {},
      openRuntimeOptions: () => {},
      openJimaku: () => {},
    }),
    cancelPendingMultiCopy: () => {
      calls.push('cancel-multi-copy');
    },
    cancelPendingMineSentenceMultiple: () => {
      calls.push('cancel-mine-sentence-multiple');
    },
  };

  assert.equal(registerOverlayShortcutsRuntime(deps), true);
  const result = unregisterOverlayShortcutsRuntime(true, deps);
  assert.equal(result, false);
  assert.deepEqual(calls, ['cancel-multi-copy', 'cancel-mine-sentence-multiple']);
});

import assert from 'node:assert/strict';
import test from 'node:test';
import type { ConfiguredShortcuts } from '../utils/shortcut-config';
import {
  registerOverlayShortcuts,
  syncOverlayShortcutsRuntime,
  unregisterOverlayShortcutsRuntime,
} from './overlay-shortcut';

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
    appendClipboardVideoToQueue: null,
    ...overrides,
  };
}

test('registerOverlayShortcuts reports active overlay shortcuts when configured', () => {
  assert.equal(
    registerOverlayShortcuts(createShortcuts({ openJimaku: 'Ctrl+J' }), {
      copySubtitle: () => {},
      copySubtitleMultiple: () => {},
      updateLastCardFromClipboard: () => {},
      triggerFieldGrouping: () => {},
      triggerSubsync: () => {},
      mineSentence: () => {},
      mineSentenceMultiple: () => {},
      toggleSecondarySub: () => {},
      markAudioCard: () => {},
      openCharacterDictionaryManager: () => {},
      openRuntimeOptions: () => {},
      openJimaku: () => {},
    }),
    true,
  );
});

test('registerOverlayShortcuts stays inactive when overlay shortcuts are absent', () => {
  assert.equal(
    registerOverlayShortcuts(createShortcuts(), {
      copySubtitle: () => {},
      copySubtitleMultiple: () => {},
      updateLastCardFromClipboard: () => {},
      triggerFieldGrouping: () => {},
      triggerSubsync: () => {},
      mineSentence: () => {},
      mineSentenceMultiple: () => {},
      toggleSecondarySub: () => {},
      markAudioCard: () => {},
      openCharacterDictionaryManager: () => {},
      openRuntimeOptions: () => {},
      openJimaku: () => {},
    }),
    false,
  );
});

test('syncOverlayShortcutsRuntime deactivates cleanly when shortcuts were active', () => {
  const calls: string[] = [];
  const result = syncOverlayShortcutsRuntime(false, true, {
    getConfiguredShortcuts: () => createShortcuts(),
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
      openCharacterDictionaryManager: () => {},
      openRuntimeOptions: () => {},
      openJimaku: () => {},
    }),
    cancelPendingMultiCopy: () => {
      calls.push('cancel-multi-copy');
    },
    cancelPendingMineSentenceMultiple: () => {
      calls.push('cancel-mine-sentence-multiple');
    },
  });

  assert.equal(result, false);
  assert.deepEqual(calls, ['cancel-multi-copy', 'cancel-mine-sentence-multiple']);
});

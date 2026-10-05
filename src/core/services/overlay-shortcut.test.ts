import assert from 'node:assert/strict';
import test from 'node:test';
import type { ConfiguredShortcuts } from '../utils/shortcut-config';
import { registerOverlayShortcuts } from './overlay-shortcut';

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
    openTsukihime: null,
    openSubtitleSelection: null,
    openSubtitleGeneration: null,
    openSessionHelp: null,
    openControllerSelect: null,
    openControllerDebug: null,
    toggleSubtitleSidebar: null,
    toggleNotificationHistory: null,
    appendClipboardVideoToQueue: null,
    ...overrides,
  };
}

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
      openTsukihime: () => {},
    }),
    false,
  );
});

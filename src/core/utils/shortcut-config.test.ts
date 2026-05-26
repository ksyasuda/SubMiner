import assert from 'node:assert/strict';
import test from 'node:test';
import type { Config } from '../../types';
import { resolveConfiguredShortcuts } from './shortcut-config';

test('forces Anki-dependent shortcuts to null when AnkiConnect is explicitly disabled', () => {
  const config: Config = {
    ankiConnect: { enabled: false },
    shortcuts: {
      copySubtitle: 'Ctrl+KeyC',
      updateLastCardFromClipboard: 'Ctrl+KeyU',
      triggerFieldGrouping: 'Alt+KeyG',
      mineSentence: 'Ctrl+Digit1',
      mineSentenceMultiple: 'Ctrl+Digit2',
      markAudioCard: 'Alt+KeyM',
    },
  };
  const defaults: Config = {
    shortcuts: {
      updateLastCardFromClipboard: 'Alt+KeyL',
      triggerFieldGrouping: 'Alt+KeyF',
      mineSentence: 'KeyQ',
      mineSentenceMultiple: 'KeyW',
      markAudioCard: 'KeyE',
    },
  };

  const resolved = resolveConfiguredShortcuts(config, defaults);

  assert.equal(resolved.updateLastCardFromClipboard, null);
  assert.equal(resolved.triggerFieldGrouping, null);
  assert.equal(resolved.mineSentence, null);
  assert.equal(resolved.mineSentenceMultiple, null);
  assert.equal(resolved.markAudioCard, null);
  assert.equal(resolved.copySubtitle, 'Ctrl+C');
});

test('keeps Anki-dependent shortcuts enabled and normalized when AnkiConnect is enabled', () => {
  const config: Config = {
    ankiConnect: { enabled: true },
    shortcuts: {
      updateLastCardFromClipboard: 'Ctrl+KeyU',
      mineSentence: 'Ctrl+Digit1',
      mineSentenceMultiple: 'Ctrl+Digit2',
    },
  };
  const defaults: Config = {
    shortcuts: {
      triggerFieldGrouping: 'Alt+KeyG',
      markAudioCard: 'Alt+KeyM',
    },
  };

  const resolved = resolveConfiguredShortcuts(config, defaults);

  assert.equal(resolved.updateLastCardFromClipboard, 'Ctrl+U');
  assert.equal(resolved.triggerFieldGrouping, 'Alt+G');
  assert.equal(resolved.mineSentence, 'Ctrl+1');
  assert.equal(resolved.mineSentenceMultiple, 'Ctrl+2');
  assert.equal(resolved.markAudioCard, 'Alt+M');
});

test('normalizes fallback shortcuts when AnkiConnect flag is unset', () => {
  const config: Config = {};
  const defaults: Config = {
    shortcuts: {
      mineSentence: 'KeyQ',
      openRuntimeOptions: 'Digit9',
      openCharacterDictionaryManager: 'Ctrl+KeyD',
    },
  };

  const resolved = resolveConfiguredShortcuts(config, defaults);

  assert.equal(resolved.mineSentence, 'Q');
  assert.equal(resolved.openRuntimeOptions, '9');
  assert.equal(resolved.openCharacterDictionaryManager, 'Ctrl+D');
});

test('preserves null shortcut overrides so defaults can be disabled', () => {
  const config: Config = {
    shortcuts: {
      copySubtitle: null,
      openJimaku: null,
      toggleSubtitleSidebar: null,
    },
  };
  const defaults: Config = {
    shortcuts: {
      copySubtitle: 'Ctrl+KeyC',
      openJimaku: 'Ctrl+Shift+KeyJ',
      toggleSubtitleSidebar: 'Backslash',
      openRuntimeOptions: 'Digit9',
    },
  };

  const resolved = resolveConfiguredShortcuts(config, defaults);

  assert.equal(resolved.copySubtitle, null);
  assert.equal(resolved.openJimaku, null);
  assert.equal(resolved.toggleSubtitleSidebar, null);
  assert.equal(resolved.openRuntimeOptions, '9');
});

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

import type { AnkiConnectConfig } from '../types';
import { KnownWordCacheManager } from './known-word-cache';

function createKnownWordCacheHarness(config: AnkiConnectConfig): {
  manager: KnownWordCacheManager;
  cleanup: () => void;
} {
  const stateDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-known-word-cache-'));
  const statePath = path.join(stateDir, 'known-words-cache.json');
  const manager = new KnownWordCacheManager({
    client: {
      findNotes: async () => [],
      notesInfo: async () => [],
    },
    getConfig: () => config,
    knownWordCacheStatePath: statePath,
    showStatusNotification: () => undefined,
  });

  return {
    manager,
    cleanup: () => {
      fs.rmSync(stateDir, { recursive: true, force: true });
    },
  };
}

test('KnownWordCacheManager invalidates persisted cache when fields.word changes', () => {
  const config: AnkiConnectConfig = {
    deck: 'Mining',
    fields: {
      word: 'Word',
    },
    knownWords: {
      highlightEnabled: true,
    },
  };
  const { manager, cleanup } = createKnownWordCacheHarness(config);

  try {
    manager.appendFromNoteInfo({
      noteId: 1,
      fields: {
        Word: { value: '猫' },
      },
    });
    assert.equal(manager.isKnownWord('猫'), true);

    config.fields = {
      ...config.fields,
      word: 'Expression',
    };

    (
      manager as unknown as {
        loadKnownWordCacheState: () => void;
      }
    ).loadKnownWordCacheState();

    assert.equal(manager.isKnownWord('猫'), false);
  } finally {
    cleanup();
  }
});

test('KnownWordCacheManager invalidates persisted cache when per-deck fields change', () => {
  const config: AnkiConnectConfig = {
    fields: {
      word: 'Word',
    },
    knownWords: {
      highlightEnabled: true,
      decks: {
        Mining: ['Word'],
      },
    },
  };
  const { manager, cleanup } = createKnownWordCacheHarness(config);

  try {
    manager.appendFromNoteInfo({
      noteId: 1,
      fields: {
        Word: { value: '猫' },
      },
    });
    assert.equal(manager.isKnownWord('猫'), true);

    config.knownWords = {
      ...config.knownWords,
      decks: {
        Mining: ['Expression'],
      },
    };

    (
      manager as unknown as {
        loadKnownWordCacheState: () => void;
      }
    ).loadKnownWordCacheState();

    assert.equal(manager.isKnownWord('猫'), false);
  } finally {
    cleanup();
  }
});

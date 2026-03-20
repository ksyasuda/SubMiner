import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

import type { AnkiConnectConfig } from '../types';
import { KnownWordCacheManager } from './known-word-cache';

function createKnownWordCacheHarness(config: AnkiConnectConfig): {
  manager: KnownWordCacheManager;
  calls: {
    findNotes: number;
    notesInfo: number;
  };
  statePath: string;
  clientState: {
    findNotesResult: number[];
    notesInfoResult: Array<{ noteId: number; fields: Record<string, { value: string }> }>;
  };
  cleanup: () => void;
} {
  const stateDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-known-word-cache-'));
  const statePath = path.join(stateDir, 'known-words-cache.json');
  const calls = {
    findNotes: 0,
    notesInfo: 0,
  };
  const clientState = {
    findNotesResult: [] as number[],
    notesInfoResult: [] as Array<{ noteId: number; fields: Record<string, { value: string }> }>,
  };
  const manager = new KnownWordCacheManager({
    client: {
      findNotes: async () => {
        calls.findNotes += 1;
        return clientState.findNotesResult;
      },
      notesInfo: async (noteIds) => {
        calls.notesInfo += 1;
        return clientState.notesInfoResult.filter((note) => noteIds.includes(note.noteId));
      },
    },
    getConfig: () => config,
    knownWordCacheStatePath: statePath,
    showStatusNotification: () => undefined,
  });

  return {
    manager,
    calls,
    statePath,
    clientState,
    cleanup: () => {
      fs.rmSync(stateDir, { recursive: true, force: true });
    },
  };
}

test('KnownWordCacheManager startLifecycle loads persisted cache without immediate rebuild', () => {
  const config: AnkiConnectConfig = {
    knownWords: {
      highlightEnabled: true,
    },
  };
  const { manager, calls, statePath, cleanup } = createKnownWordCacheHarness(config);

  try {
    fs.writeFileSync(
      statePath,
      JSON.stringify({
        version: 2,
        refreshedAtMs: Date.now(),
        scope: '{"refreshMinutes":1440,"scope":"is:note","fieldsWord":""}',
        words: ['猫'],
        notes: {
          '1': ['猫'],
        },
      }),
      'utf-8',
    );

    manager.startLifecycle();

    assert.equal(manager.isKnownWord('猫'), true);
    assert.equal(calls.findNotes, 0);
    assert.equal(calls.notesInfo, 0);
  } finally {
    manager.stopLifecycle();
    cleanup();
  }
});

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

test('KnownWordCacheManager refresh incrementally reconciles deleted and edited note words', async () => {
  const config: AnkiConnectConfig = {
    fields: {
      word: 'Word',
    },
    knownWords: {
      highlightEnabled: true,
    },
  };
  const { manager, statePath, clientState, cleanup } = createKnownWordCacheHarness(config);

  try {
    fs.writeFileSync(
      statePath,
      JSON.stringify({
        version: 2,
        refreshedAtMs: 1,
        scope: '{"refreshMinutes":1440,"scope":"is:note","fieldsWord":"Word"}',
        words: ['猫', '犬'],
        notes: {
          '1': ['猫'],
          '2': ['犬'],
        },
      }),
      'utf-8',
    );

    (
      manager as unknown as {
        loadKnownWordCacheState: () => void;
      }
    ).loadKnownWordCacheState();

    clientState.findNotesResult = [1];
    clientState.notesInfoResult = [
      {
        noteId: 1,
        fields: {
          Word: { value: '鳥' },
        },
      },
    ];

    await manager.refresh(true);

    assert.equal(manager.isKnownWord('猫'), false);
    assert.equal(manager.isKnownWord('犬'), false);
    assert.equal(manager.isKnownWord('鳥'), true);

    const persisted = JSON.parse(fs.readFileSync(statePath, 'utf-8')) as {
      version: number;
      words: string[];
      notes?: Record<string, string[]>;
    };
    assert.equal(persisted.version, 2);
    assert.deepEqual(persisted.words.sort(), ['鳥']);
    assert.deepEqual(persisted.notes, {
      '1': ['鳥'],
    });
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

test('KnownWordCacheManager skips immediate append when addMinedWordsImmediately is disabled', () => {
  const config: AnkiConnectConfig = {
    knownWords: {
      highlightEnabled: true,
      addMinedWordsImmediately: false,
    },
  };
  const { manager, statePath, cleanup } = createKnownWordCacheHarness(config);

  try {
    manager.appendFromNoteInfo({
      noteId: 1,
      fields: {
        Expression: { value: '猫' },
      },
    });

    assert.equal(manager.isKnownWord('猫'), false);
    assert.equal(fs.existsSync(statePath), false);
  } finally {
    cleanup();
  }
});

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

import type { AnkiConnectConfig } from '../types/anki';
import { KnownWordCacheManager } from './known-word-cache';

async function waitForCondition(
  condition: () => boolean,
  timeoutMs = 500,
  intervalMs = 10,
): Promise<void> {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    if (condition()) {
      return;
    }
    await new Promise((resolve) => setTimeout(resolve, intervalMs));
  }
  throw new Error('Timed out waiting for condition');
}

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
    findNotesByQuery: Map<string, number[]>;
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
    findNotesByQuery: new Map<string, number[]>(),
  };
  const manager = new KnownWordCacheManager({
    client: {
      findNotes: async (query) => {
        calls.findNotes += 1;
        if (clientState.findNotesByQuery.has(query)) {
          return clientState.findNotesByQuery.get(query) ?? [];
        }
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

test('KnownWordCacheManager startLifecycle keeps fresh persisted cache without immediate refresh', async () => {
  const config: AnkiConnectConfig = {
    knownWords: {
      highlightEnabled: true,
      refreshMinutes: 60,
    },
  };
  const { manager, calls, statePath, cleanup } = createKnownWordCacheHarness(config);

  try {
    fs.writeFileSync(
      statePath,
      JSON.stringify({
        version: 2,
        refreshedAtMs: Date.now(),
        scope: '{"refreshMinutes":60,"scope":"is:note","fieldsWord":""}',
        words: ['猫'],
        notes: {
          '1': ['猫'],
        },
      }),
      'utf-8',
    );

    manager.startLifecycle();
    await new Promise((resolve) => setTimeout(resolve, 25));

    assert.equal(manager.isKnownWord('猫'), true);
    assert.equal(calls.findNotes, 0);
    assert.equal(calls.notesInfo, 0);
  } finally {
    manager.stopLifecycle();
    cleanup();
  }
});

test('KnownWordCacheManager startLifecycle immediately refreshes stale persisted cache', async () => {
  const config: AnkiConnectConfig = {
    fields: {
      word: 'Word',
    },
    knownWords: {
      highlightEnabled: true,
      refreshMinutes: 1,
    },
  };
  const { manager, calls, statePath, clientState, cleanup } = createKnownWordCacheHarness(config);

  try {
    fs.writeFileSync(
      statePath,
      JSON.stringify({
        version: 2,
        refreshedAtMs: Date.now() - 61_000,
        scope: '{"refreshMinutes":1,"scope":"is:note","fieldsWord":"Word"}',
        words: ['猫'],
        notes: {
          '1': ['猫'],
        },
      }),
      'utf-8',
    );

    clientState.findNotesResult = [1];
    clientState.notesInfoResult = [
      {
        noteId: 1,
        fields: {
          Word: { value: '犬' },
        },
      },
    ];

    manager.startLifecycle();
    await waitForCondition(() => calls.findNotes === 1 && calls.notesInfo === 1);

    assert.equal(manager.isKnownWord('猫'), false);
    assert.equal(manager.isKnownWord('犬'), true);
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

test('KnownWordCacheManager skips malformed note info without fields', async () => {
  const config: AnkiConnectConfig = {
    fields: {
      word: 'Word',
    },
    knownWords: {
      highlightEnabled: true,
    },
  };
  const { manager, clientState, cleanup } = createKnownWordCacheHarness(config);

  try {
    clientState.findNotesResult = [1, 2];
    clientState.notesInfoResult = [
      {
        noteId: 1,
        fields: undefined as unknown as Record<string, { value: string }>,
      },
      {
        noteId: 2,
        fields: {
          Word: { value: '猫' },
        },
      },
    ];

    await manager.refresh(true);

    assert.equal(manager.isKnownWord('猫'), true);
    assert.equal(manager.isKnownWord('犬'), false);
  } finally {
    cleanup();
  }
});

test('KnownWordCacheManager preserves cache state key captured before refresh work', async () => {
  const config: AnkiConnectConfig = {
    fields: {
      word: 'Word',
    },
    knownWords: {
      highlightEnabled: true,
      refreshMinutes: 1,
    },
  };
  const stateDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-known-word-cache-key-'));
  const statePath = path.join(stateDir, 'known-words-cache.json');
  let notesInfoStarted = false;
  let releaseNotesInfo!: () => void;
  const notesInfoGate = new Promise<void>((resolve) => {
    releaseNotesInfo = resolve;
  });
  const manager = new KnownWordCacheManager({
    client: {
      findNotes: async () => [1],
      notesInfo: async () => {
        notesInfoStarted = true;
        await notesInfoGate;
        return [
          {
            noteId: 1,
            fields: {
              Word: { value: '猫' },
            },
          },
        ];
      },
    },
    getConfig: () => config,
    knownWordCacheStatePath: statePath,
    showStatusNotification: () => undefined,
  });

  try {
    const refreshPromise = manager.refresh(true);
    await waitForCondition(() => notesInfoStarted);

    config.fields = {
      ...config.fields,
      word: 'Expression',
    };
    releaseNotesInfo();
    await refreshPromise;

    const persisted = JSON.parse(fs.readFileSync(statePath, 'utf-8')) as {
      scope: string;
      words: string[];
    };
    assert.equal(persisted.scope, '{"refreshMinutes":1,"scope":"is:note","fieldsWord":"Word"}');
    assert.deepEqual(persisted.words, ['猫']);
  } finally {
    fs.rmSync(stateDir, { recursive: true, force: true });
  }
});

test('KnownWordCacheManager does not borrow fields from other decks during refresh', async () => {
  const config: AnkiConnectConfig = {
    knownWords: {
      highlightEnabled: true,
      decks: {
        Mining: [],
        Reading: ['AltWord'],
      },
    },
  };
  const { manager, clientState, cleanup } = createKnownWordCacheHarness(config);

  try {
    clientState.findNotesByQuery.set('deck:"Mining"', [1]);
    clientState.findNotesByQuery.set('deck:"Reading"', []);
    clientState.notesInfoResult = [
      {
        noteId: 1,
        fields: {
          AltWord: { value: '猫' },
        },
      },
    ];

    await manager.refresh(true);

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

test('KnownWordCacheManager preserves deck-specific field mappings during refresh', async () => {
  const config: AnkiConnectConfig = {
    knownWords: {
      highlightEnabled: true,
      decks: {
        Mining: ['Expression'],
        Reading: ['Word'],
      },
    },
  };
  const { manager, clientState, cleanup } = createKnownWordCacheHarness(config);

  try {
    clientState.findNotesByQuery.set('deck:"Mining"', [1]);
    clientState.findNotesByQuery.set('deck:"Reading"', [2]);
    clientState.notesInfoResult = [
      {
        noteId: 1,
        fields: {
          Expression: { value: '猫' },
          Word: { value: 'should-not-count' },
        },
      },
      {
        noteId: 2,
        fields: {
          Word: { value: '犬' },
          Expression: { value: 'also-ignored' },
        },
      },
    ];

    await manager.refresh(true);

    assert.equal(manager.isKnownWord('猫'), true);
    assert.equal(manager.isKnownWord('犬'), true);
    assert.equal(manager.isKnownWord('should-not-count'), false);
    assert.equal(manager.isKnownWord('also-ignored'), false);
  } finally {
    cleanup();
  }
});

test('KnownWordCacheManager uses the current deck fields for immediate append', () => {
  const config: AnkiConnectConfig = {
    deck: 'Mining',
    fields: {
      word: 'Word',
    },
    knownWords: {
      highlightEnabled: true,
      decks: {
        Mining: ['Expression'],
        Reading: ['Word'],
      },
    },
  };
  const { manager, cleanup } = createKnownWordCacheHarness(config);

  try {
    manager.appendFromNoteInfo({
      noteId: 1,
      fields: {
        Expression: { value: '猫' },
        Word: { value: 'should-not-count' },
      },
    });

    assert.equal(manager.isKnownWord('猫'), true);
    assert.equal(manager.isKnownWord('should-not-count'), false);
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

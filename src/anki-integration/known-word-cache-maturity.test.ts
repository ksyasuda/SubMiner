import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

import type { AnkiConnectConfig } from '../types/anki';
import { KnownWordCacheManager, getKnownWordCacheLifecycleConfig } from './known-word-cache';

interface HarnessNoteInfo {
  noteId: number;
  fields: Record<string, { value: string }>;
}

function createMaturityHarness(config: AnkiConnectConfig): {
  manager: KnownWordCacheManager;
  calls: { findNotes: number; notesInfo: number; queries: string[] };
  statePath: string;
  clientState: {
    findNotesResult: number[];
    notesInfoResult: HarnessNoteInfo[];
    findNotesByQuery: Map<string, number[]>;
    failedQueries: Set<string>;
  };
  createSiblingManager: () => KnownWordCacheManager;
  cleanup: () => void;
} {
  const stateDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-known-word-maturity-'));
  const statePath = path.join(stateDir, 'known-words-cache.json');
  const calls = { findNotes: 0, notesInfo: 0, queries: [] as string[] };
  const clientState = {
    findNotesResult: [] as number[],
    notesInfoResult: [] as HarnessNoteInfo[],
    findNotesByQuery: new Map<string, number[]>(),
    failedQueries: new Set<string>(),
  };
  const deps = {
    client: {
      findNotes: async (query: string) => {
        calls.findNotes += 1;
        calls.queries.push(query);
        if (clientState.failedQueries.has(query)) {
          throw new Error(`Anki unavailable for query: ${query}`);
        }
        if (clientState.findNotesByQuery.has(query)) {
          return clientState.findNotesByQuery.get(query) ?? [];
        }
        return clientState.findNotesResult;
      },
      notesInfo: async (noteIds: number[]) => {
        calls.notesInfo += 1;
        return clientState.notesInfoResult.filter((note) => noteIds.includes(note.noteId));
      },
    },
    getConfig: () => config,
    knownWordCacheStatePath: statePath,
    showStatusNotification: () => undefined,
  };
  return {
    manager: new KnownWordCacheManager(deps),
    calls,
    statePath,
    clientState,
    createSiblingManager: () => new KnownWordCacheManager(deps),
    cleanup: () => {
      fs.rmSync(stateDir, { recursive: true, force: true });
    },
  };
}

function maturityConfig(overrides: Partial<AnkiConnectConfig> = {}): AnkiConnectConfig {
  return {
    deck: 'Mining',
    fields: { word: 'Word' },
    knownWords: {
      highlightEnabled: true,
      maturityEnabled: true,
      refreshMinutes: 60,
    },
    ...overrides,
  };
}

test('lifecycle config key is unchanged when maturity is disabled', () => {
  const disabled: AnkiConnectConfig = {
    knownWords: { highlightEnabled: true, refreshMinutes: 60 },
  };
  // Upgrading users keep their persisted cache: the key must not gain fields
  // while maturity is off.
  assert.equal(
    getKnownWordCacheLifecycleConfig(disabled),
    '{"refreshMinutes":60,"scope":"all","fieldsWord":""}',
  );

  const enabled: AnkiConnectConfig = {
    knownWords: { highlightEnabled: true, maturityEnabled: true, refreshMinutes: 60 },
  };
  assert.equal(
    getKnownWordCacheLifecycleConfig(enabled),
    '{"refreshMinutes":60,"scope":"all","fieldsWord":"","maturity":21}',
  );

  const customThreshold: AnkiConnectConfig = {
    knownWords: {
      highlightEnabled: true,
      maturityEnabled: true,
      matureThresholdDays: 30,
      refreshMinutes: 60,
    },
  };
  assert.equal(
    getKnownWordCacheLifecycleConfig(customThreshold),
    '{"refreshMinutes":60,"scope":"all","fieldsWord":"","maturity":30}',
  );
});

test('refresh fetches tier sets and getKnownWordTier classifies notes', async () => {
  const { manager, calls, clientState, cleanup } = createMaturityHarness(maturityConfig());

  try {
    clientState.findNotesByQuery.set('deck:"Mining"', [1, 2, 3, 4]);
    clientState.findNotesByQuery.set('deck:"Mining" prop:ivl>=21', [1]);
    clientState.findNotesByQuery.set('deck:"Mining" prop:ivl>=1 prop:ivl<21', [2]);
    clientState.findNotesByQuery.set('deck:"Mining" is:learn', [3]);
    clientState.notesInfoResult = [
      { noteId: 1, fields: { Word: { value: '猫' } } },
      { noteId: 2, fields: { Word: { value: '犬' } } },
      { noteId: 3, fields: { Word: { value: '鳥' } } },
      { noteId: 4, fields: { Word: { value: '魚' } } },
    ];

    await manager.refresh(true);

    assert.equal(calls.findNotes, 4);
    assert.equal(manager.getKnownWordTier('猫'), 'mature');
    assert.equal(manager.getKnownWordTier('犬'), 'young');
    assert.equal(manager.getKnownWordTier('鳥'), 'learning');
    assert.equal(manager.getKnownWordTier('魚'), 'new');
    assert.equal(manager.getKnownWordTier('馬'), null);
    // Boolean matching still works alongside tiers.
    assert.equal(manager.isKnownWord('猫'), true);
    assert.equal(manager.isKnownWord('魚'), true);
  } finally {
    cleanup();
  }
});

test('a note with cards in several tiers counts as its most mature card', async () => {
  const { manager, clientState, cleanup } = createMaturityHarness(maturityConfig());

  try {
    clientState.findNotesByQuery.set('deck:"Mining"', [1]);
    clientState.findNotesByQuery.set('deck:"Mining" prop:ivl>=21', [1]);
    clientState.findNotesByQuery.set('deck:"Mining" prop:ivl>=1 prop:ivl<21', [1]);
    clientState.findNotesByQuery.set('deck:"Mining" is:learn', [1]);
    clientState.notesInfoResult = [{ noteId: 1, fields: { Word: { value: '猫' } } }];

    await manager.refresh(true);

    assert.equal(manager.getKnownWordTier('猫'), 'mature');
  } finally {
    cleanup();
  }
});

test('a word matched by several notes takes the most mature note tier', async () => {
  const { manager, clientState, cleanup } = createMaturityHarness(maturityConfig());

  try {
    clientState.findNotesByQuery.set('deck:"Mining"', [1, 2]);
    clientState.findNotesByQuery.set('deck:"Mining" prop:ivl>=21', []);
    clientState.findNotesByQuery.set('deck:"Mining" prop:ivl>=1 prop:ivl<21', [2]);
    clientState.findNotesByQuery.set('deck:"Mining" is:learn', [1]);
    clientState.notesInfoResult = [
      { noteId: 1, fields: { Word: { value: '猫' } } },
      { noteId: 2, fields: { Word: { value: '猫' } } },
    ];

    await manager.refresh(true);

    assert.equal(manager.getKnownWordTier('猫'), 'young');
  } finally {
    cleanup();
  }
});

test('tiers are reading-aware for words with several readings', async () => {
  const { manager, clientState, cleanup } = createMaturityHarness(maturityConfig());

  try {
    clientState.findNotesByQuery.set('deck:"Mining"', [1, 2]);
    clientState.findNotesByQuery.set('deck:"Mining" prop:ivl>=21', [1]);
    clientState.findNotesByQuery.set('deck:"Mining" prop:ivl>=1 prop:ivl<21', []);
    clientState.findNotesByQuery.set('deck:"Mining" is:learn', [2]);
    clientState.notesInfoResult = [
      { noteId: 1, fields: { Word: { value: '床' }, Reading: { value: 'とこ' } } },
      { noteId: 2, fields: { Word: { value: '床' }, Reading: { value: 'ゆか' } } },
    ];

    await manager.refresh(true);

    assert.equal(manager.getKnownWordTier('床', 'とこ'), 'mature');
    assert.equal(manager.getKnownWordTier('床', 'ゆか'), 'learning');
    // No reading given: fail-open across readings, most mature wins.
    assert.equal(manager.getKnownWordTier('床'), 'mature');
    // Unknown reading for a reading-locked word: no match, no tier.
    assert.equal(manager.getKnownWordTier('床', 'しょう'), null);
  } finally {
    cleanup();
  }
});

test('reading-only fallback resolves tiers unless opted out', async () => {
  const { manager, clientState, cleanup } = createMaturityHarness(maturityConfig());

  try {
    clientState.findNotesByQuery.set('deck:"Mining"', [1]);
    clientState.findNotesByQuery.set('deck:"Mining" prop:ivl>=21', [1]);
    clientState.findNotesByQuery.set('deck:"Mining" prop:ivl>=1 prop:ivl<21', []);
    clientState.findNotesByQuery.set('deck:"Mining" is:learn', []);
    clientState.notesInfoResult = [
      { noteId: 1, fields: { Word: { value: '警告' }, Reading: { value: 'けいこく' } } },
    ];

    await manager.refresh(true);

    assert.equal(manager.getKnownWordTier('けいこく'), 'mature');
    assert.equal(
      manager.getKnownWordTier('けいこく', undefined, { allowReadingOnlyMatch: false }),
      null,
    );
  } finally {
    cleanup();
  }
});

test('getKnownWordTier returns null and skips tier queries when maturity is disabled', async () => {
  const config = maturityConfig();
  config.knownWords = { ...config.knownWords, maturityEnabled: false };
  const { manager, calls, clientState, cleanup } = createMaturityHarness(config);

  try {
    clientState.findNotesByQuery.set('deck:"Mining"', [1]);
    clientState.notesInfoResult = [{ noteId: 1, fields: { Word: { value: '猫' } } }];

    await manager.refresh(true);

    assert.equal(calls.findNotes, 1);
    assert.equal(manager.isKnownWord('猫'), true);
    assert.equal(manager.getKnownWordTier('猫'), null);
  } finally {
    cleanup();
  }
});

test('refresh preserves known-word cache when maturity lookup fails', async () => {
  const { manager, statePath, clientState, cleanup } = createMaturityHarness(maturityConfig());

  try {
    clientState.findNotesByQuery.set('deck:"Mining"', [1]);
    clientState.failedQueries.add('deck:"Mining" prop:ivl>=21');
    clientState.notesInfoResult = [{ noteId: 1, fields: { Word: { value: '猫' } } }];

    await manager.refresh(true);

    assert.equal(manager.isKnownWord('猫'), true);
    assert.equal(manager.getKnownWordTier('猫'), null);
    const persisted = JSON.parse(fs.readFileSync(statePath, 'utf-8')) as {
      version: number;
      tiers: Record<string, string>;
    };
    assert.equal(persisted.version, 4);
    assert.deepEqual(persisted.tiers, {});
  } finally {
    cleanup();
  }
});

test('tiers persist to v4 state and reload without refetching', async () => {
  const { manager, calls, statePath, clientState, createSiblingManager, cleanup } =
    createMaturityHarness(maturityConfig());
  const originalDateNow = Date.now;

  try {
    Date.now = () => 120_000;
    clientState.findNotesByQuery.set('deck:"Mining"', [1, 2]);
    clientState.findNotesByQuery.set('deck:"Mining" prop:ivl>=21', [1]);
    clientState.findNotesByQuery.set('deck:"Mining" prop:ivl>=1 prop:ivl<21', []);
    clientState.findNotesByQuery.set('deck:"Mining" is:learn', [2]);
    clientState.notesInfoResult = [
      { noteId: 1, fields: { Word: { value: '猫' } } },
      { noteId: 2, fields: { Word: { value: '犬' } } },
    ];

    await manager.refresh(true);

    const persisted = JSON.parse(fs.readFileSync(statePath, 'utf-8')) as {
      version: number;
      tiers: Record<string, string>;
    };
    assert.equal(persisted.version, 4);
    assert.deepEqual(persisted.tiers, { '1': 'mature', '2': 'learning' });

    const callsBeforeReload = calls.findNotes;
    const reloaded = createSiblingManager();
    reloaded.startLifecycle();
    try {
      assert.equal(reloaded.getKnownWordTier('猫'), 'mature');
      assert.equal(reloaded.getKnownWordTier('犬'), 'learning');
      assert.equal(calls.findNotes, callsBeforeReload);
    } finally {
      reloaded.stopLifecycle();
    }
  } finally {
    Date.now = originalDateNow;
    cleanup();
  }
});

test('appendFromNoteInfo marks freshly mined notes as new tier', async () => {
  const { manager, cleanup } = createMaturityHarness(maturityConfig());

  try {
    manager.appendFromNoteInfo({
      noteId: 7,
      fields: { Word: { value: '猫' } },
    });

    assert.equal(manager.isKnownWord('猫'), true);
    assert.equal(manager.getKnownWordTier('猫'), 'new');
  } finally {
    cleanup();
  }
});

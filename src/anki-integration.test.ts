import test from 'node:test';
import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AnkiIntegration } from './anki-integration';
import { FieldGroupingMergeCollaborator } from './anki-integration/field-grouping-merge';
import { AnkiConnectConfig } from './types';

interface IntegrationTestContext {
  integration: AnkiIntegration;
  calls: {
    findNotes: number;
    notesInfo: number;
  };
  stateDir: string;
}

function createIntegrationTestContext(
  options: {
    highlightEnabled?: boolean;
    onFindNotes?: () => Promise<number[]>;
    onNotesInfo?: () => Promise<unknown[]>;
    stateDirPrefix?: string;
  } = {},
): IntegrationTestContext {
  const calls = {
    findNotes: 0,
    notesInfo: 0,
  };

  const stateDir = fs.mkdtempSync(
    path.join(os.tmpdir(), options.stateDirPrefix ?? 'subminer-anki-integration-'),
  );
  const knownWordCacheStatePath = path.join(stateDir, 'known-words-cache.json');

  const client = {
    findNotes: async () => {
      calls.findNotes += 1;
      if (options.onFindNotes) {
        return options.onFindNotes();
      }
      return [] as number[];
    },
    notesInfo: async () => {
      calls.notesInfo += 1;
      if (options.onNotesInfo) {
        return options.onNotesInfo();
      }
      return [] as unknown[];
    },
  } as {
    findNotes: () => Promise<number[]>;
    notesInfo: () => Promise<unknown[]>;
  };

  const integration = new AnkiIntegration(
    {
      nPlusOne: {
        highlightEnabled: options.highlightEnabled ?? true,
      },
    },
    {} as never,
    {} as never,
    undefined,
    undefined,
    undefined,
    knownWordCacheStatePath,
  );

  const integrationWithClient = integration as unknown as {
    client: {
      findNotes: () => Promise<number[]>;
      notesInfo: () => Promise<unknown[]>;
    };
  };
  integrationWithClient.client = client;

  const privateState = integration as unknown as {
    knownWordsScope: string;
    knownWordsLastRefreshedAtMs: number;
  };
  privateState.knownWordsScope = 'is:note';
  privateState.knownWordsLastRefreshedAtMs = Date.now();

  return {
    integration,
    calls,
    stateDir,
  };
}

function cleanupIntegrationTestContext(ctx: IntegrationTestContext): void {
  fs.rmSync(ctx.stateDir, { recursive: true, force: true });
}

function resolveFieldName(availableFieldNames: string[], preferredName: string): string | null {
  const exact = availableFieldNames.find((name) => name === preferredName);
  if (exact) return exact;

  const lower = preferredName.toLowerCase();
  return availableFieldNames.find((name) => name.toLowerCase() === lower) ?? null;
}

function createFieldGroupingMergeCollaborator(options?: {
  config?: Partial<AnkiConnectConfig>;
  currentSubtitleText?: string;
  generatedMedia?: {
    audioField?: string;
    audioValue?: string;
    imageField?: string;
    imageValue?: string;
    miscInfoValue?: string;
  };
}): FieldGroupingMergeCollaborator {
  const config = {
    fields: {
      sentence: 'Sentence',
      audio: 'ExpressionAudio',
      image: 'Picture',
      ...(options?.config?.fields ?? {}),
    },
    ...(options?.config ?? {}),
  } as AnkiConnectConfig;

  return new FieldGroupingMergeCollaborator({
    getConfig: () => config,
    getEffectiveSentenceCardConfig: () => ({
      sentenceField: 'Sentence',
      audioField: 'SentenceAudio',
    }),
    getCurrentSubtitleText: () => options?.currentSubtitleText,
    resolveFieldName,
    resolveNoteFieldName: (noteInfo, preferredName) => {
      if (!preferredName) return null;
      return resolveFieldName(Object.keys(noteInfo.fields), preferredName);
    },
    extractFields: (fields) => {
      const result: Record<string, string> = {};
      for (const [key, value] of Object.entries(fields)) {
        result[key.toLowerCase()] = value.value || '';
      }
      return result;
    },
    processSentence: (mpvSentence) => `${mpvSentence}::processed`,
    generateMediaForMerge: async () => options?.generatedMedia ?? {},
    warnFieldParseOnce: () => undefined,
  });
}

test('AnkiIntegration.refreshKnownWordCache bypasses stale checks', async () => {
  const ctx = createIntegrationTestContext();

  try {
    await ctx.integration.refreshKnownWordCache();

    assert.equal(ctx.calls.findNotes, 1);
    assert.equal(ctx.calls.notesInfo, 0);
  } finally {
    cleanupIntegrationTestContext(ctx);
  }
});

test('AnkiIntegration.refreshKnownWordCache skips work when highlight mode is disabled', async () => {
  const ctx = createIntegrationTestContext({
    highlightEnabled: false,
    stateDirPrefix: 'subminer-anki-integration-disabled-',
  });

  try {
    await ctx.integration.refreshKnownWordCache();

    assert.equal(ctx.calls.findNotes, 0);
    assert.equal(ctx.calls.notesInfo, 0);
  } finally {
    cleanupIntegrationTestContext(ctx);
  }
});

test('AnkiIntegration.refreshKnownWordCache deduplicates concurrent refreshes', async () => {
  let releaseFindNotes: (() => void) | undefined;
  const findNotesPromise = new Promise<void>((resolve) => {
    releaseFindNotes = resolve;
  });

  const ctx = createIntegrationTestContext({
    onFindNotes: async () => {
      await findNotesPromise;
      return [] as number[];
    },
    stateDirPrefix: 'subminer-anki-integration-concurrent-',
  });

  const first = ctx.integration.refreshKnownWordCache();
  await Promise.resolve();
  const second = ctx.integration.refreshKnownWordCache();

  if (releaseFindNotes !== undefined) {
    releaseFindNotes();
  }

  await Promise.all([first, second]);

  try {
    assert.equal(ctx.calls.findNotes, 1);
    assert.equal(ctx.calls.notesInfo, 0);
  } finally {
    cleanupIntegrationTestContext(ctx);
  }
});

test('AnkiIntegration does not allocate proxy server when proxy transport is disabled', () => {
  const integration = new AnkiIntegration(
    {
      enabled: true,
      proxy: {
        enabled: false,
      },
    } as never,
    {} as never,
    {} as never,
  );

  const privateState = integration as unknown as {
    proxyServer: unknown | null;
  };
  assert.equal(privateState.proxyServer, null);
});

test('FieldGroupingMergeCollaborator synchronizes ExpressionAudio from merged SentenceAudio', async () => {
  const collaborator = createFieldGroupingMergeCollaborator();

  const merged = await collaborator.computeFieldGroupingMergedFields(
    101,
    202,
    {
      noteId: 101,
      fields: {
        SentenceAudio: { value: '[sound:keep.mp3]' },
        ExpressionAudio: { value: '[sound:stale.mp3]' },
      },
    },
    {
      noteId: 202,
      fields: {
        SentenceAudio: { value: '[sound:new.mp3]' },
      },
    },
    false,
  );

  assert.equal(
    merged.SentenceAudio,
    '<span data-group-id="101">[sound:keep.mp3]</span><span data-group-id="202">[sound:new.mp3]</span>',
  );
  assert.equal(merged.ExpressionAudio, merged.SentenceAudio);
});

test('FieldGroupingMergeCollaborator uses generated media fallback when source lacks audio', async () => {
  const collaborator = createFieldGroupingMergeCollaborator({
    generatedMedia: {
      audioField: 'SentenceAudio',
      audioValue: '[sound:generated.mp3]',
    },
  });

  const merged = await collaborator.computeFieldGroupingMergedFields(
    11,
    22,
    {
      noteId: 11,
      fields: {
        SentenceAudio: { value: '' },
      },
    },
    {
      noteId: 22,
      fields: {
        SentenceAudio: { value: '' },
      },
    },
    true,
  );

  assert.equal(merged.SentenceAudio, '<span data-group-id="22">[sound:generated.mp3]</span>');
});

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
      knownWords: {
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

test('AnkiIntegration notifies when mined note info updates known words', () => {
  const ctx = createIntegrationTestContext({
    stateDirPrefix: 'subminer-anki-integration-known-update-',
  });
  let notifications = 0;

  try {
    const integrationState = ctx.integration as unknown as {
      config: AnkiConnectConfig;
      appendKnownWordsFromNoteInfo: (noteInfo: {
        noteId: number;
        fields: Record<string, { value: string }>;
      }) => void;
    };
    integrationState.config.deck = 'Mining';
    integrationState.config.knownWords = {
      ...integrationState.config.knownWords,
      decks: {
        Mining: ['Word'],
      },
    };
    ctx.integration.setKnownWordCacheUpdatedCallback(() => {
      notifications += 1;
    });
    integrationState.appendKnownWordsFromNoteInfo({
      noteId: 42,
      fields: {
        Word: { value: '食べる' },
      },
    });

    assert.equal(ctx.integration.isKnownWord('食べる'), true);
    assert.equal(notifications, 1);
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

test('AnkiIntegration resolves merged-away note ids to the kept note id', () => {
  const ctx = createIntegrationTestContext({
    stateDirPrefix: 'subminer-anki-integration-note-redirect-',
  });

  try {
    const integrationWithInternals = ctx.integration as unknown as {
      rememberMergedNoteIds: (deletedNoteId: number, keptNoteId: number) => void;
    };
    integrationWithInternals.rememberMergedNoteIds(111, 222);
    integrationWithInternals.rememberMergedNoteIds(222, 333);

    assert.equal(ctx.integration.resolveCurrentNoteId(111), 333);
    assert.equal(ctx.integration.resolveCurrentNoteId(222), 333);
    assert.equal(ctx.integration.resolveCurrentNoteId(333), 333);
    assert.equal(ctx.integration.resolveCurrentNoteId(444), 444);
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
    runtime: {
      proxyServer: unknown | null;
    };
  };
  assert.equal(privateState.runtime.proxyServer, null);
});

test('AnkiIntegration triggers field grouping after a local duplicate sentence card is created', async () => {
  const integration = new AnkiIntegration(
    {
      isKiku: {
        enabled: true,
        fieldGrouping: 'manual',
      },
    } as never,
    {} as never,
    {} as never,
  );

  let groupingTriggered = 0;
  const internals = integration as unknown as {
    cardCreationService: {
      createSentenceCard: (
        sentence: string,
        startTime: number,
        endTime: number,
        secondarySubText?: string,
      ) => Promise<boolean>;
    };
    fieldGroupingService: {
      triggerFieldGroupingForLastAddedCard: () => Promise<void>;
    };
  };
  internals.cardCreationService = {
    createSentenceCard: async () => {
      integration.trackDuplicateNoteIdsForNote(42, [7]);
      return true;
    },
  };
  internals.fieldGroupingService = {
    triggerFieldGroupingForLastAddedCard: async () => {
      groupingTriggered += 1;
    },
  };

  assert.equal(await integration.createSentenceCard('duplicate sentence', 0, 1), true);
  assert.equal(groupingTriggered, 1);
});

test('AnkiIntegration marks partial update notifications as failures in OSD mode', async () => {
  const osdMessages: string[] = [];
  const integration = new AnkiIntegration(
    {
      behavior: {
        notificationType: 'osd',
      },
    },
    {} as never,
    {} as never,
    (text) => {
      osdMessages.push(text);
    },
  );

  await (
    integration as unknown as {
      showNotification: (
        noteId: number,
        label: string | number,
        errorSuffix?: string,
      ) => Promise<void>;
    }
  ).showNotification(42, 'taberu', 'image failed');

  assert.deepEqual(osdMessages, ['x Updated card: taberu (image failed)']);
});

test('FieldGroupingMergeCollaborator keeps SentenceAudio grouped without overwriting ExpressionAudio', async () => {
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
    '<span data-group-id="202">[sound:new.mp3]</span><span data-group-id="101">[sound:keep.mp3]</span>',
  );
  assert.equal('ExpressionAudio' in merged, false);
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

test('FieldGroupingMergeCollaborator keeps independent groups for identical sentence, audio, and image values', async () => {
  const collaborator = createFieldGroupingMergeCollaborator();

  const merged = await collaborator.computeFieldGroupingMergedFields(
    202,
    101,
    {
      noteId: 202,
      fields: {
        Sentence: { value: 'same sentence' },
        SentenceAudio: { value: '[sound:same.mp3]' },
        Picture: { value: '<img src="same.png">' },
        ExpressionAudio: { value: '[sound:same.mp3]' },
      },
    },
    {
      noteId: 101,
      fields: {
        Sentence: { value: 'same sentence' },
        SentenceAudio: { value: '[sound:same.mp3]' },
        Picture: { value: '<img src="same.png">' },
      },
    },
    false,
  );

  assert.equal(
    merged.Sentence,
    '<span data-group-id="202">same sentence</span><span data-group-id="101">same sentence</span>',
  );
  assert.equal(
    merged.SentenceAudio,
    '<span data-group-id="202">[sound:same.mp3]</span><span data-group-id="101">[sound:same.mp3]</span>',
  );
  assert.equal(
    merged.Picture,
    '<img data-group-id="202" src="same.png"><img data-group-id="101" src="same.png">',
  );
  assert.equal('ExpressionAudio' in merged, false);
});

test('AnkiIntegration.formatMiscInfoPattern avoids leaking Jellyfin api_key query params', () => {
  const integration = new AnkiIntegration(
    {
      metadata: {
        pattern: '[SubMiner] %f (%t)',
      },
    } as never,
    {} as never,
    {
      currentSubText: '',
      currentVideoPath:
        'stream?static=true&api_key=secret-token&MediaSourceId=a762ab23d26d4347e3cacdb83aaae405&AudioStreamIndex=3',
      currentTimePos: 426,
      currentSubStart: 426,
      currentSubEnd: 428,
      currentAudioStreamIndex: 3,
      currentMediaTitle: '[Jellyfin/direct] Bocchi the Rock! - S01E02',
      send: () => true,
    } as unknown as never,
  );

  const privateApi = integration as unknown as {
    formatMiscInfoPattern: (fallbackFilename: string, startTimeSeconds?: number) => string;
  };
  const result = privateApi.formatMiscInfoPattern('audio_123.mp3', 426);

  assert.equal(result, '[SubMiner] [Jellyfin/direct] Bocchi the Rock! - S01E02 (00:07:06)');
  assert.equal(result.includes('api_key='), false);
});

import test from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import { once } from 'node:events';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AnkiIntegration } from './anki-integration';
import { FieldGroupingMergeCollaborator } from './anki-integration/field-grouping-merge';
import type { MediaInput } from './media-input';
import { AnkiConnectConfig } from './types';

type TestOverlayNotificationPayload = {
  id?: string;
  title: string;
  body?: string;
  image?: string;
  variant?: string;
  persistent?: boolean;
  actions?: Array<{ id: string; label: string; noteId?: number }>;
};

interface IntegrationTestContext {
  integration: AnkiIntegration;
  calls: {
    findNotes: number;
    notesInfo: number;
  };
  stateDir: string;
}

function describeMediaInputForTest(input: MediaInput): string {
  if (typeof input === 'string') {
    return input;
  }
  return `${input.path}:${input.source ?? 'raw'}`;
}

function createIntegrationTestContext(
  options: {
    highlightEnabled?: boolean;
    nPlusOneEnabled?: boolean;
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
      nPlusOne:
        options.nPlusOneEnabled === undefined
          ? undefined
          : {
              enabled: options.nPlusOneEnabled,
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
  privateState.knownWordsScope = 'all';
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

test('AnkiIntegration.refreshKnownWordCache notifies annotation cache listeners', async () => {
  const ctx = createIntegrationTestContext({
    stateDirPrefix: 'subminer-anki-integration-refresh-notify-',
  });
  let notifications = 0;

  try {
    ctx.integration.setKnownWordCacheUpdatedCallback(() => {
      notifications += 1;
    });

    await ctx.integration.refreshKnownWordCache();

    assert.equal(notifications, 1);
  } finally {
    cleanupIntegrationTestContext(ctx);
  }
});

test('AnkiIntegration.refreshKnownWordCache notifies when n+1 is enabled without highlights', async () => {
  const ctx = createIntegrationTestContext({
    highlightEnabled: false,
    nPlusOneEnabled: true,
    stateDirPrefix: 'subminer-anki-integration-nplusone-notify-',
  });
  let notifications = 0;

  try {
    ctx.integration.setKnownWordCacheUpdatedCallback(() => {
      notifications += 1;
    });

    await ctx.integration.refreshKnownWordCache();

    assert.equal(ctx.calls.findNotes, 1);
    assert.equal(notifications, 1);
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

function processSentenceWithConfig(
  config: Partial<AnkiConnectConfig>,
  mpvSentence: string,
  noteFields: Record<string, string>,
): string {
  const integration = new AnkiIntegration(config as AnkiConnectConfig, {} as never, {} as never);
  return (
    integration as unknown as {
      processSentence: (sentence: string, fields: Record<string, string>) => string;
    }
  ).processSentence(mpvSentence, noteFields);
}

function processSentenceFuriganaWithConfig(
  config: Partial<AnkiConnectConfig>,
  sentenceFurigana: string,
  noteFields: Record<string, string>,
): string {
  const integration = new AnkiIntegration(config as AnkiConnectConfig, {} as never, {} as never);
  return (
    integration as unknown as {
      processSentenceFurigana: (sentence: string, fields: Record<string, string>) => string;
    }
  ).processSentenceFurigana(sentenceFurigana, noteFields);
}

test('AnkiIntegration highlights mined word from expression field when sentence has no bold marker', () => {
  const processed = processSentenceWithConfig(
    {
      fields: {
        word: 'Expression',
        sentence: 'Sentence',
      },
      behavior: {
        highlightWord: true,
      },
    },
    '先日 貴様らが潜入した キールダンジョンから―',
    {
      expression: '潜入',
      sentence: '先日 貴様らが潜入した キールダンジョンから―',
    },
  );

  assert.equal(processed, '先日 貴様らが<b>潜入</b>した キールダンジョンから―');
});

test('AnkiIntegration keeps existing Yomitan bold target when present', () => {
  const processed = processSentenceWithConfig(
    {
      fields: {
        word: 'Expression',
        sentence: 'Sentence',
      },
      behavior: {
        highlightWord: true,
      },
    },
    '先日 貴様らが潜入した キールダンジョンから―',
    {
      expression: '潜入',
      sentence: '<b>潜入した</b>',
    },
  );

  assert.equal(processed, '先日 貴様らが<b>潜入した</b> キールダンジョンから―');
});

test('AnkiIntegration leaves sentence plain when word highlighting is disabled', () => {
  const processed = processSentenceWithConfig(
    {
      fields: {
        word: 'Expression',
        sentence: 'Sentence',
      },
      behavior: {
        highlightWord: false,
      },
    },
    '先日 貴様らが潜入した キールダンジョンから―',
    {
      expression: '潜入',
      sentence: '<b>潜入</b>',
    },
  );

  assert.equal(processed, '先日 貴様らが潜入した キールダンジョンから―');
});

test('AnkiIntegration highlights mined word in sentence furigana field', () => {
  const processed = processSentenceFuriganaWithConfig(
    {
      fields: {
        word: 'Expression',
        sentence: 'Sentence',
      },
      behavior: {
        highlightWord: true,
      },
    },
    '<span class="term"><ruby>不思議<rt>ふしぎ</rt></ruby></span><span class="term">な</span><span class="term"><ruby>特技<rt>とくぎ</rt></ruby></span><span class="term">を</span>',
    {
      expression: '特技',
      sentence: '不思議な特技を',
    },
  );

  assert.equal(
    processed,
    '<span class="term"><ruby>不思議<rt>ふしぎ</rt></ruby></span><span class="term">な</span><b><span class="term"><ruby>特技<rt>とくぎ</rt></ruby></span></b><span class="term">を</span>',
  );
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

test('AnkiIntegration reports an occupied proxy address through its notification seam', async () => {
  const occupiedServer = http.createServer();
  occupiedServer.listen(0, '127.0.0.1');
  await once(occupiedServer, 'listening');
  const occupiedAddress = occupiedServer.address();
  assert.ok(occupiedAddress && typeof occupiedAddress === 'object');
  const stateDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-anki-proxy-collision-'));
  const overlayNotifications: TestOverlayNotificationPayload[] = [];
  const integration = new AnkiIntegration(
    {
      enabled: true,
      url: 'http://127.0.0.1:8765',
      proxy: {
        enabled: true,
        host: '127.0.0.1',
        port: occupiedAddress.port,
        upstreamUrl: 'http://127.0.0.1:8765',
      },
      behavior: {
        notificationType: 'overlay',
      },
      knownWords: {
        highlightEnabled: false,
      },
      nPlusOne: {
        enabled: false,
      },
    } as never,
    {} as never,
    {} as never,
    undefined,
    undefined,
    undefined,
    path.join(stateDir, 'known-words-cache.json'),
    {},
    undefined,
    (payload) => {
      overlayNotifications.push(payload as TestOverlayNotificationPayload);
    },
  );

  try {
    integration.start();
    await integration.waitUntilReady();

    assert.deepEqual(overlayNotifications, [
      {
        title: 'SubMiner',
        body: `AnkiConnect proxy unavailable because http://127.0.0.1:${occupiedAddress.port} is already in use. Change ankiConnect.proxy.port or stop the process using that address.`,
        variant: 'info',
      },
    ]);
  } finally {
    integration.stop();
    occupiedServer.close();
    await once(occupiedServer, 'close');
    fs.rmSync(stateDir, { recursive: true, force: true });
  }
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

test('AnkiIntegration applies ready YouTube cache media to every queued note id', async () => {
  const osdMessages: string[] = [];
  const updatedNotes: Array<{ noteId: number; fields: Record<string, string> }> = [];
  const storedMedia: string[] = [];
  const mediaInputs: string[] = [];

  const integration = new AnkiIntegration(
    {
      fields: {
        image: 'Picture',
      },
      media: {
        imageFormat: 'jpg',
      },
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

  const internals = integration as unknown as {
    client: {
      notesInfo: (noteIds: number[]) => Promise<unknown[]>;
      updateNoteFields: (noteId: number, fields: Record<string, string>) => Promise<void>;
      storeMediaFile: (filename: string, data: Buffer) => Promise<void>;
    };
    mediaGenerator: {
      generateAudio: (
        path: MediaInput,
        startTime: number,
        endTime: number,
        audioPadding?: number,
        audioStreamIndex?: number,
        normalizeAudio?: boolean,
        volumeScale?: number,
      ) => Promise<Buffer>;
      generateScreenshot: (path: MediaInput) => Promise<Buffer>;
    };
    queuePendingYoutubeMediaUpdate: (job: {
      sourceUrl: string;
      noteId: number;
      startTime: number;
      endTime: number;
      label: string | number;
      audioStreamIndex?: number;
      audioFieldName?: string;
      imageFieldName?: string;
      generateAudio: boolean;
      generateImage: boolean;
      volumeScale?: number;
    }) => void;
  };
  internals.client = {
    notesInfo: async (noteIds) =>
      noteIds.map((noteId) => ({
        noteId,
        fields: {
          SentenceAudio: { value: '' },
          Picture: { value: '' },
        },
      })),
    updateNoteFields: async (noteId, fields) => {
      updatedNotes.push({ noteId, fields });
    },
    storeMediaFile: async (filename) => {
      storedMedia.push(filename);
    },
  };
  internals.mediaGenerator = {
    generateAudio: async (
      mediaPath,
      _startTime,
      _endTime,
      _audioPadding,
      audioStreamIndex,
      _normalizeAudio,
      volumeScale,
    ) => {
      mediaInputs.push(
        `audio:${describeMediaInputForTest(mediaPath)}:${audioStreamIndex ?? 'auto'}:${volumeScale}`,
      );
      return Buffer.from('audio');
    },
    generateScreenshot: async (mediaPath) => {
      mediaInputs.push(`image:${describeMediaInputForTest(mediaPath)}`);
      return Buffer.from('image');
    },
  };
  internals.queuePendingYoutubeMediaUpdate({
    sourceUrl: 'https://www.youtube.com/watch?v=abc123',
    noteId: 101,
    startTime: 10,
    endTime: 12,
    label: 'first',
    audioStreamIndex: 22,
    audioFieldName: 'SentenceAudio',
    imageFieldName: 'Picture',
    generateAudio: true,
    generateImage: true,
    volumeScale: 0.25,
  });
  internals.queuePendingYoutubeMediaUpdate({
    sourceUrl: 'https://youtu.be/abc123',
    noteId: 202,
    startTime: 20,
    endTime: 22,
    label: 'second',
    audioStreamIndex: 23,
    audioFieldName: 'SentenceAudio',
    imageFieldName: 'Picture',
    generateAudio: true,
    generateImage: true,
    volumeScale: 0.8,
  });

  await integration.handleYoutubeMediaCacheReady('https://youtu.be/abc123', '/tmp/media.mkv');

  assert.deepEqual(mediaInputs, [
    'audio:/tmp/media.mkv:youtube-cache:auto:0.25',
    'image:/tmp/media.mkv:youtube-cache',
    'audio:/tmp/media.mkv:youtube-cache:auto:0.8',
    'image:/tmp/media.mkv:youtube-cache',
  ]);
  assert.deepEqual(
    updatedNotes.map((update) => update.noteId),
    [101, 202],
  );
  const firstUpdate = updatedNotes[0];
  const secondUpdate = updatedNotes[1];
  assert.ok(firstUpdate);
  assert.ok(secondUpdate);
  assert.match(firstUpdate.fields.SentenceAudio ?? '', /^\[sound:audio_/);
  assert.match(firstUpdate.fields.Picture ?? '', /^<img src="image_/);
  assert.match(secondUpdate.fields.SentenceAudio ?? '', /^\[sound:audio_/);
  assert.match(secondUpdate.fields.Picture ?? '', /^<img src="image_/);
  assert.equal(storedMedia.length, 4);
  assert.equal(
    osdMessages.some((message) =>
      message.includes('YouTube media cache ready. Adding media to 2 queued cards.'),
    ),
    true,
  );
});

test('AnkiIntegration reports partial queued YouTube media updates separately from failures', async () => {
  const osdMessages: string[] = [];
  const updatedNotes: Array<{ noteId: number; fields: Record<string, string> }> = [];
  const notifications: Array<{ noteId: number; label: string | number; suffix?: string }> = [];

  const integration = new AnkiIntegration(
    {
      fields: {
        image: 'Picture',
      },
      media: {
        imageFormat: 'jpg',
      },
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

  const internals = integration as unknown as {
    client: {
      notesInfo: (noteIds: number[]) => Promise<unknown[]>;
      updateNoteFields: (noteId: number, fields: Record<string, string>) => Promise<void>;
      storeMediaFile: () => Promise<void>;
    };
    mediaGenerator: {
      generateAudio: () => Promise<Buffer>;
      generateScreenshot: () => Promise<Buffer>;
    };
    queuePendingYoutubeMediaUpdate: (job: {
      sourceUrl: string;
      noteId: number;
      startTime: number;
      endTime: number;
      label: string | number;
      audioFieldName?: string;
      imageFieldName?: string;
      generateAudio: boolean;
      generateImage: boolean;
    }) => void;
    showNotification: (noteId: number, label: string | number, suffix?: string) => Promise<void>;
  };
  internals.client = {
    notesInfo: async (noteIds) =>
      noteIds.map((noteId) => ({
        noteId,
        fields: {
          SentenceAudio: { value: '' },
          Picture: { value: '' },
        },
      })),
    updateNoteFields: async (noteId, fields) => {
      updatedNotes.push({ noteId, fields });
    },
    storeMediaFile: async () => undefined,
  };
  internals.mediaGenerator = {
    generateAudio: async () => {
      throw new Error('audio stream not found');
    },
    generateScreenshot: async () => Buffer.from('image'),
  };
  internals.showNotification = async (noteId, label, suffix) => {
    notifications.push({ noteId, label, suffix });
  };

  internals.queuePendingYoutubeMediaUpdate({
    sourceUrl: 'https://www.youtube.com/watch?v=partial',
    noteId: 303,
    startTime: 10,
    endTime: 12,
    label: 'partial',
    audioFieldName: 'SentenceAudio',
    imageFieldName: 'Picture',
    generateAudio: true,
    generateImage: true,
  });

  await integration.handleYoutubeMediaCacheReady('https://youtu.be/partial', '/tmp/media.mkv');

  assert.equal(updatedNotes.length, 1);
  assert.match(updatedNotes[0]?.fields.Picture ?? '', /^<img src="image_/);
  assert.equal(updatedNotes[0]?.fields.SentenceAudio, undefined);
  assert.deepEqual(notifications, [{ noteId: 303, label: 'partial', suffix: 'audio failed' }]);
  assert.equal(
    osdMessages.some((message) =>
      message.includes('Queued YouTube media finished with 0 updated, 1 partial, and 0 failed.'),
    ),
    true,
  );
});

test('AnkiIntegration queues YouTube media updates against recovered source URLs', async () => {
  const updatedNotes: Array<{ noteId: number; fields: Record<string, string> }> = [];
  const storedMedia: string[] = [];
  const audioVolumeScales: Array<number | undefined> = [];
  let mpvVolume = 30;

  const integration = new AnkiIntegration(
    {
      fields: {
        image: 'Picture',
      },
      media: {
        imageFormat: 'jpg',
      },
    },
    {} as never,
    {
      currentVideoPath: 'https://rr1---sn.example.googlevideo.com/videoplayback?expire=1777777777',
      currentSubStart: 10,
      currentSubEnd: 12,
      currentTimePos: 11,
      requestProperty: async (name: string) => {
        assert.equal(name, 'volume');
        return mpvVolume;
      },
    } as never,
    () => undefined,
    undefined,
    undefined,
    undefined,
    {},
    undefined,
    undefined,
    async () => null,
    () => true,
    () => 'https://www.youtube.com/watch?v=abc123',
  );

  const internals = integration as unknown as {
    client: {
      notesInfo: (noteIds: number[]) => Promise<unknown[]>;
      updateNoteFields: (noteId: number, fields: Record<string, string>) => Promise<void>;
      storeMediaFile: (filename: string) => Promise<void>;
    };
    mediaGenerator: {
      generateAudio: (
        path: MediaInput,
        startTime: number,
        endTime: number,
        audioPadding?: number,
        audioStreamIndex?: number,
        normalizeAudio?: boolean,
        volumeScale?: number,
      ) => Promise<Buffer>;
      generateScreenshot: () => Promise<Buffer>;
    };
    queuePendingYoutubeMediaUpdateForNote: (job: {
      noteId: number;
      noteInfo: { noteId: number; fields: Record<string, { value: string }> };
      label: string | number;
    }) => Promise<boolean>;
    showNotification: () => Promise<void>;
  };
  internals.client = {
    notesInfo: async (noteIds) =>
      noteIds.map((noteId) => ({
        noteId,
        fields: {
          SentenceAudio: { value: '' },
          Picture: { value: '' },
        },
      })),
    updateNoteFields: async (noteId, fields) => {
      updatedNotes.push({ noteId, fields });
    },
    storeMediaFile: async (filename) => {
      storedMedia.push(filename);
    },
  };
  internals.mediaGenerator = {
    generateAudio: async (
      _path,
      _startTime,
      _endTime,
      _audioPadding,
      _audioStreamIndex,
      _normalizeAudio,
      volumeScale,
    ) => {
      audioVolumeScales.push(volumeScale);
      return Buffer.from('audio');
    },
    generateScreenshot: async () => Buffer.from('image'),
  };
  internals.showNotification = async () => undefined;

  const queued = await internals.queuePendingYoutubeMediaUpdateForNote({
    noteId: 404,
    noteInfo: {
      noteId: 404,
      fields: {
        SentenceAudio: { value: '' },
        Picture: { value: '' },
      },
    },
    label: 'resolved source',
  });
  mpvVolume = 90;
  await integration.handleYoutubeMediaCacheReady('https://youtu.be/abc123', '/tmp/media.mkv');

  assert.equal(queued, true);
  assert.equal(updatedNotes.length, 1);
  assert.equal(updatedNotes[0]?.noteId, 404);
  assert.match(updatedNotes[0]?.fields.SentenceAudio ?? '', /^\[sound:audio_/);
  assert.match(updatedNotes[0]?.fields.Picture ?? '', /^<img src="image_/);
  assert.equal(storedMedia.length, 2);
  assert.deepEqual(audioVolumeScales, [0.3 ** 3]);
});

test('AnkiIntegration passes audio normalization config for ready cached YouTube audio', async () => {
  const audioCalls: Array<{
    path: string;
    audioStreamIndex?: number;
    normalizeAudio?: boolean;
    volumeScale?: number;
  }> = [];
  const requestedProperties: string[] = [];

  const integration = new AnkiIntegration(
    {
      media: {
        audioPadding: 0,
        normalizeAudio: false,
      },
    },
    {} as never,
    {
      currentVideoPath: 'https://www.youtube.com/watch?v=abc123',
      currentAudioStreamIndex: 0,
      currentSubStart: 10,
      currentSubEnd: 12,
      currentTimePos: 11,
      requestProperty: async (name: string) => {
        requestedProperties.push(name);
        return 55;
      },
    } as never,
    () => undefined,
    undefined,
    undefined,
    undefined,
    {},
    undefined,
    undefined,
    async () => '/tmp/subminer-youtube-media-cache/media.mkv',
    () => true,
  );

  const internals = integration as unknown as {
    mediaGenerator: {
      generateAudio: (
        path: { path: string },
        startTime: number,
        endTime: number,
        audioPadding?: number,
        audioStreamIndex?: number,
        normalizeAudio?: boolean,
        volumeScale?: number,
      ) => Promise<Buffer>;
    };
    generateAudio: () => Promise<Buffer | null>;
  };
  internals.mediaGenerator = {
    generateAudio: async (
      path,
      _startTime,
      _endTime,
      _audioPadding,
      audioStreamIndex,
      normalizeAudio,
      volumeScale,
    ) => {
      audioCalls.push({ path: path.path, audioStreamIndex, normalizeAudio, volumeScale });
      return Buffer.from('audio');
    },
  };

  await internals.generateAudio();

  assert.deepEqual(requestedProperties, ['volume']);
  assert.deepEqual(audioCalls, [
    {
      path: '/tmp/subminer-youtube-media-cache/media.mkv',
      audioStreamIndex: undefined,
      normalizeAudio: false,
      volumeScale: 0.55 ** 3,
    },
  ]);
});

test('AnkiIntegration announces ready YouTube cache when no queued notes exist', async () => {
  const overlayNotifications: TestOverlayNotificationPayload[] = [];

  const integration = new AnkiIntegration(
    {
      behavior: {
        notificationType: 'overlay',
      },
    },
    {} as never,
    {} as never,
    undefined,
    undefined,
    undefined,
    undefined,
    {},
    undefined,
    (payload) => {
      overlayNotifications.push(payload as TestOverlayNotificationPayload);
    },
  );

  await integration.handleYoutubeMediaCacheReady('https://youtu.be/abc123', '/tmp/media.mkv');

  assert.equal(overlayNotifications.length, 1);
  assert.equal(overlayNotifications[0]?.title, 'SubMiner');
  assert.equal(overlayNotifications[0]?.body, 'YouTube media cache ready.');
});

test('AnkiIntegration can let caller own no-queued YouTube cache ready notification', async () => {
  const overlayNotifications: TestOverlayNotificationPayload[] = [];

  const integration = new AnkiIntegration(
    {
      behavior: {
        notificationType: 'overlay',
      },
    },
    {} as never,
    {} as never,
    undefined,
    undefined,
    undefined,
    undefined,
    {},
    undefined,
    (payload) => {
      overlayNotifications.push(payload as TestOverlayNotificationPayload);
    },
  );

  await integration.handleYoutubeMediaCacheReady('https://youtu.be/abc123', '/tmp/media.mkv', {
    notifyNoQueued: false,
  });

  assert.deepEqual(overlayNotifications, []);
});

test('AnkiIntegration embeds generated notification image on overlay mined-card notifications', async () => {
  const desktopNotifications: Array<{ title: string; body?: string; icon?: string }> = [];
  const overlayNotifications: TestOverlayNotificationPayload[] = [];
  const generatedFrom: Array<{ videoPath: string; timestamp: number }> = [];
  const cleanupPaths: string[] = [];
  const notificationIconPath = path.join(os.tmpdir(), 'subminer-notification-icon.png');

  const integration = new AnkiIntegration(
    {
      behavior: {
        notificationType: 'both',
      },
    },
    {} as never,
    {
      currentVideoPath: '/tmp/show.mkv',
      currentTimePos: 123.45,
    } as never,
    undefined,
    (title, options) => {
      desktopNotifications.push({ title, body: options.body, icon: options.icon });
    },
    undefined,
    undefined,
    {},
    undefined,
    (payload) => {
      overlayNotifications.push(payload as TestOverlayNotificationPayload);
    },
  );

  (
    integration as unknown as {
      mediaGenerator: {
        generateNotificationIcon: (videoPath: string, timestamp: number) => Promise<Buffer>;
        writeNotificationIconToFile: (iconBuffer: Buffer, noteId: number) => string;
        scheduleNotificationIconCleanup: (filePath: string) => void;
      };
    }
  ).mediaGenerator = {
    generateNotificationIcon: async (videoPath, timestamp) => {
      generatedFrom.push({ videoPath, timestamp });
      return Buffer.from('png');
    },
    writeNotificationIconToFile: (iconBuffer, noteId) => {
      assert.equal(iconBuffer.toString(), 'png');
      assert.equal(noteId, 42);
      return notificationIconPath;
    },
    scheduleNotificationIconCleanup: (filePath) => {
      cleanupPaths.push(filePath);
    },
  };

  await (
    integration as unknown as {
      showNotification: (noteId: number, label: string | number) => Promise<void>;
    }
  ).showNotification(42, '食べる');

  assert.deepEqual(generatedFrom, [{ videoPath: '/tmp/show.mkv', timestamp: 123.45 }]);
  assert.equal(overlayNotifications.length, 1);
  assert.equal(overlayNotifications[0]?.title, 'Anki Card Updated');
  assert.equal(overlayNotifications[0]?.body, 'Updated card: 食べる');
  assert.equal(
    overlayNotifications[0]?.image,
    `data:image/png;base64,${Buffer.from('png').toString('base64')}`,
  );
  assert.deepEqual(overlayNotifications[0]?.actions, [
    { id: 'open-anki-card', label: 'Open in Anki', noteId: 42 },
  ]);
  assert.deepEqual(desktopNotifications, [
    {
      title: 'Anki Card Updated',
      body: 'Updated card: 食べる',
      icon: notificationIconPath,
    },
  ]);
  assert.deepEqual(cleanupPaths, [notificationIconPath]);
});

test('AnkiIntegration keeps overlay card-update progress visible until the terminal notification', async () => {
  const overlayNotifications: TestOverlayNotificationPayload[] = [];
  const integration = new AnkiIntegration(
    {
      behavior: {
        notificationType: 'overlay',
      },
    },
    {} as never,
    {} as never,
    undefined,
    undefined,
    undefined,
    undefined,
    {},
    undefined,
    (payload) => {
      overlayNotifications.push(payload);
    },
  );
  const updateNotifications = integration as unknown as {
    beginUpdateProgress: (message: string) => void;
    showNotification: (noteId: number, label: string | number) => Promise<void>;
  };

  updateNotifications.beginUpdateProgress('Updating card');
  await updateNotifications.showNotification(42, '食べる');

  assert.deepEqual(
    overlayNotifications.map(({ id, variant, persistent }) => ({ id, variant, persistent })),
    [
      { id: 'anki-update-progress', variant: 'progress', persistent: true },
      { id: 'anki-update-progress', variant: 'success', persistent: false },
    ],
  );
});

test('AnkiIntegration dismisses persistent overlay update progress when no terminal notification replaces it', () => {
  const overlayNotifications: TestOverlayNotificationPayload[] = [];
  const dismissedIds: string[] = [];
  const integration = new AnkiIntegration(
    {
      behavior: {
        notificationType: 'overlay',
      },
    },
    {} as never,
    {} as never,
    undefined,
    undefined,
    undefined,
    undefined,
    {},
    undefined,
    (payload) => {
      overlayNotifications.push(payload);
    },
    undefined,
    undefined,
    undefined,
    (id) => {
      dismissedIds.push(id);
    },
  );
  const updateNotifications = integration as unknown as {
    beginUpdateProgress: (message: string) => void;
    endUpdateProgress: () => void;
  };

  updateNotifications.beginUpdateProgress('Updating card');
  updateNotifications.endUpdateProgress();

  assert.equal(overlayNotifications[0]?.persistent, true);
  assert.deepEqual(dismissedIds, ['anki-update-progress']);
});

test('AnkiIntegration keeps overlay notification image when temp icon write fails', async () => {
  const desktopNotifications: Array<{ title: string; body?: string; icon?: string }> = [];
  const overlayNotifications: TestOverlayNotificationPayload[] = [];
  const cleanupPaths: string[] = [];

  const integration = new AnkiIntegration(
    {
      behavior: {
        notificationType: 'both',
      },
    },
    {} as never,
    {
      currentVideoPath: '/tmp/show.mkv',
      currentTimePos: 123.45,
    } as never,
    undefined,
    (title, options) => {
      desktopNotifications.push({ title, body: options.body, icon: options.icon });
    },
    undefined,
    undefined,
    {},
    undefined,
    (payload) => {
      overlayNotifications.push(payload as TestOverlayNotificationPayload);
    },
  );

  (
    integration as unknown as {
      mediaGenerator: {
        generateNotificationIcon: () => Promise<Buffer>;
        writeNotificationIconToFile: () => string;
        scheduleNotificationIconCleanup: (filePath: string) => void;
      };
    }
  ).mediaGenerator = {
    generateNotificationIcon: async () => Buffer.from('png'),
    writeNotificationIconToFile: () => {
      throw new Error('disk full');
    },
    scheduleNotificationIconCleanup: (filePath) => {
      cleanupPaths.push(filePath);
    },
  };

  await (
    integration as unknown as {
      showNotification: (noteId: number, label: string | number) => Promise<void>;
    }
  ).showNotification(42, '食べる');

  assert.equal(
    overlayNotifications[0]?.image,
    `data:image/png;base64,${Buffer.from('png').toString('base64')}`,
  );
  assert.deepEqual(desktopNotifications, [
    {
      title: 'Anki Card Updated',
      body: 'Updated card: 食べる',
      icon: undefined,
    },
  ]);
  assert.deepEqual(cleanupPaths, []);
});

test('AnkiIntegration routes workflow status notifications through configured surfaces', async () => {
  const osdMessages: string[] = [];
  const desktopMessages: string[] = [];
  const overlayMessages: string[] = [];
  const integration = new AnkiIntegration(
    {
      behavior: {
        notificationType: 'both',
      },
    },
    {} as never,
    {} as never,
    (text) => {
      osdMessages.push(text);
    },
    (title, options) => {
      desktopMessages.push(`${title}:${options.body ?? ''}`);
    },
    undefined,
    undefined,
    {},
    undefined,
    (payload) => {
      overlayMessages.push(`${payload.title}:${payload.body ?? ''}:${payload.variant ?? ''}`);
    },
  );

  assert.equal(await integration.createSentenceCard('食べる', 0, 1), false);

  assert.deepEqual(osdMessages, []);
  assert.deepEqual(overlayMessages, ['SubMiner:No video loaded:info']);
  assert.deepEqual(desktopMessages, ['SubMiner:No video loaded']);
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

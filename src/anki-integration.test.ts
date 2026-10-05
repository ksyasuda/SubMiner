import test from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import { once } from 'node:events';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AnkiIntegration } from './anki-integration';
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

type GenerateAudioFn = (
  path: MediaInput,
  startTime: number,
  endTime: number,
  audioPadding?: number,
  audioStreamIndex?: number,
  normalizeAudio?: boolean,
  volumeScale?: number,
) => Promise<Buffer>;

type PendingYoutubeJob = {
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
};

/**
 * The private members these tests swap or call. AnkiIntegration has no constructor seam for its
 * AnkiConnect client or media generator, so tests replace them after construction.
 */
type IntegrationInternals = {
  client: {
    findNotes?: () => Promise<number[]>;
    notesInfo?: (noteIds: number[]) => Promise<unknown[]>;
    updateNoteFields?: (noteId: number, fields: Record<string, string>) => Promise<void>;
    storeMediaFile?: (filename: string, data: Buffer) => Promise<void>;
  };
  mediaGenerator: {
    generateAudio?: GenerateAudioFn;
    generateScreenshot?: (path: MediaInput) => Promise<Buffer>;
    generateNotificationIcon?: (videoPath: string, timestamp: number) => Promise<Buffer>;
    writeNotificationIconToFile?: (iconBuffer: Buffer, noteId: number) => string;
    scheduleNotificationIconCleanup?: (filePath: string) => void;
  };
  config: AnkiConnectConfig;
  cardCreationService: { createSentenceCard: () => Promise<boolean> };
  fieldGroupingService: { triggerFieldGroupingForLastAddedCard: () => Promise<void> };
  showNotification: (noteId: number, label: string | number, errorSuffix?: string) => Promise<void>;
  beginUpdateProgress: (message: string) => void;
  endUpdateProgress: () => void;
  processSentence: (sentence: string, fields: Record<string, string>) => string;
  processSentenceFurigana: (sentence: string, fields: Record<string, string>) => string;
  queuePendingYoutubeMediaUpdate: (job: PendingYoutubeJob) => void;
  queuePendingYoutubeMediaUpdateForNote: (job: {
    noteId: number;
    noteInfo: { noteId: number; fields: Record<string, { value: string }> };
    label: string | number;
  }) => Promise<boolean>;
  generateAudio: () => Promise<Buffer | null>;
  formatMiscInfoPattern: (fallbackFilename: string, startTimeSeconds?: number) => string;
  appendKnownWordsFromNoteInfo: (noteInfo: {
    noteId: number;
    fields: Record<string, { value: string }>;
  }) => void;
  rememberMergedNoteIds: (deletedNoteId: number, keptNoteId: number) => void;
};

/** Named stand-ins for AnkiIntegration's long positional constructor arguments. */
type IntegrationHooks = {
  mpvClient?: Record<string, unknown>;
  osd?: (text: string) => void;
  desktop?: (title: string, options: { body?: string; icon?: string }) => void;
  overlay?: (payload: TestOverlayNotificationPayload) => void;
  dismissOverlay?: (id: string) => void;
  knownWordCacheStatePath?: string;
  getCachedMediaPath?: () => Promise<string | null>;
  shouldRequireRemoteMediaCache?: () => boolean;
  getYoutubeMediaSourceUrl?: () => string | null;
};

function createIntegration(
  config: Partial<AnkiConnectConfig>,
  hooks: IntegrationHooks = {},
): AnkiIntegration {
  return new AnkiIntegration(
    config as AnkiConnectConfig,
    {} as never,
    (hooks.mpvClient ?? {}) as never,
    hooks.osd,
    hooks.desktop as never,
    undefined,
    hooks.knownWordCacheStatePath,
    {},
    undefined,
    hooks.overlay as never,
    hooks.getCachedMediaPath,
    hooks.shouldRequireRemoteMediaCache,
    hooks.getYoutubeMediaSourceUrl,
    hooks.dismissOverlay,
  );
}

function internalsOf(integration: AnkiIntegration): IntegrationInternals {
  return integration as unknown as IntegrationInternals;
}

function describeMediaInputForTest(input: MediaInput): string {
  if (typeof input === 'string') {
    return input;
  }
  return `${input.path}:${input.source ?? 'raw'}`;
}

// --- Known-word cache -------------------------------------------------------

type KnownWordsOptions = {
  highlightEnabled?: boolean;
  nPlusOneEnabled?: boolean;
  onFindNotes?: () => Promise<number[]>;
};

type KnownWordsContext = {
  integration: AnkiIntegration;
  calls: { findNotes: number; notesInfo: number };
};

/** Runs `run` against an integration whose AnkiConnect client is a counting fake; cleans up its cache dir. */
async function withKnownWordsIntegration(
  options: KnownWordsOptions,
  run: (ctx: KnownWordsContext) => Promise<void> | void,
): Promise<void> {
  const calls = { findNotes: 0, notesInfo: 0 };
  const stateDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-anki-integration-'));

  try {
    const integration = createIntegration(
      {
        knownWords: { highlightEnabled: options.highlightEnabled ?? true },
        nPlusOne:
          options.nPlusOneEnabled === undefined ? undefined : { enabled: options.nPlusOneEnabled },
      },
      { knownWordCacheStatePath: path.join(stateDir, 'known-words-cache.json') },
    );
    internalsOf(integration).client = {
      findNotes: async () => {
        calls.findNotes += 1;
        return options.onFindNotes ? options.onFindNotes() : [];
      },
      notesInfo: async () => {
        calls.notesInfo += 1;
        return [];
      },
    };
    await run({ integration, calls });
  } finally {
    fs.rmSync(stateDir, { recursive: true, force: true });
  }
}

const KNOWN_WORD_REFRESH_CASES: Array<{
  name: string;
  options: KnownWordsOptions;
  expected: { findNotes: number; notifications: number };
}> = [
  {
    name: 'refreshes once, skipping stale checks, and notifies annotation cache listeners',
    options: {},
    expected: { findNotes: 1, notifications: 1 },
  },
  {
    name: 'notifies when n+1 is enabled without highlights',
    options: { highlightEnabled: false, nPlusOneEnabled: true },
    expected: { findNotes: 1, notifications: 1 },
  },
  {
    name: 'skips work when highlight mode is disabled',
    options: { highlightEnabled: false },
    expected: { findNotes: 0, notifications: 0 },
  },
];

for (const c of KNOWN_WORD_REFRESH_CASES) {
  test(`AnkiIntegration.refreshKnownWordCache ${c.name}`, async () => {
    await withKnownWordsIntegration(c.options, async ({ integration, calls }) => {
      let notifications = 0;
      integration.setKnownWordCacheUpdatedCallback(() => {
        notifications += 1;
      });

      await integration.refreshKnownWordCache();

      assert.equal(calls.findNotes, c.expected.findNotes);
      assert.equal(calls.notesInfo, 0);
      assert.equal(notifications, c.expected.notifications);
    });
  });
}

test('AnkiIntegration.refreshKnownWordCache deduplicates concurrent refreshes', async () => {
  let releaseFindNotes: () => void = () => {};
  const findNotesGate = new Promise<void>((resolve) => {
    releaseFindNotes = resolve;
  });

  await withKnownWordsIntegration(
    {
      onFindNotes: async () => {
        await findNotesGate;
        return [];
      },
    },
    async ({ integration, calls }) => {
      const first = integration.refreshKnownWordCache();
      await Promise.resolve();
      const second = integration.refreshKnownWordCache();
      releaseFindNotes();
      await Promise.all([first, second]);

      assert.equal(calls.findNotes, 1);
      assert.equal(calls.notesInfo, 0);
    },
  );
});

test('AnkiIntegration notifies when mined note info updates known words', async () => {
  await withKnownWordsIntegration({}, ({ integration }) => {
    let notifications = 0;
    const internals = internalsOf(integration);
    internals.config.deck = 'Mining';
    internals.config.knownWords = {
      ...internals.config.knownWords,
      decks: { Mining: ['Word'] },
    };
    integration.setKnownWordCacheUpdatedCallback(() => {
      notifications += 1;
    });

    internals.appendKnownWordsFromNoteInfo({ noteId: 42, fields: { Word: { value: '食べる' } } });

    assert.equal(integration.isKnownWord('食べる'), true);
    assert.equal(notifications, 1);
  });
});

test('AnkiIntegration resolves merged-away note ids to the kept note id', () => {
  const integration = createIntegration({});
  const internals = internalsOf(integration);
  internals.rememberMergedNoteIds(111, 222);
  internals.rememberMergedNoteIds(222, 333);

  assert.equal(integration.resolveCurrentNoteId(111), 333);
  assert.equal(integration.resolveCurrentNoteId(222), 333);
  assert.equal(integration.resolveCurrentNoteId(333), 333);
  assert.equal(integration.resolveCurrentNoteId(444), 444);
});

// --- Sentence highlighting --------------------------------------------------

const MPV_SENTENCE = '先日 貴様らが潜入した キールダンジョンから―';

const HIGHLIGHT_CASES = [
  {
    name: 'highlights mined word from expression field when sentence has no bold marker',
    highlightWord: true,
    noteSentence: MPV_SENTENCE,
    expected: '先日 貴様らが<b>潜入</b>した キールダンジョンから―',
  },
  {
    name: 'keeps existing Yomitan bold target when present',
    highlightWord: true,
    noteSentence: '<b>潜入した</b>',
    expected: '先日 貴様らが<b>潜入した</b> キールダンジョンから―',
  },
  {
    name: 'leaves sentence plain when word highlighting is disabled',
    highlightWord: false,
    noteSentence: '<b>潜入</b>',
    expected: MPV_SENTENCE,
  },
];

for (const c of HIGHLIGHT_CASES) {
  test(`AnkiIntegration ${c.name}`, () => {
    const integration = createIntegration({
      fields: { word: 'Expression', sentence: 'Sentence' },
      behavior: { highlightWord: c.highlightWord },
    });

    const processed = internalsOf(integration).processSentence(MPV_SENTENCE, {
      expression: '潜入',
      sentence: c.noteSentence,
    });

    assert.equal(processed, c.expected);
  });
}

test('AnkiIntegration highlights mined word in sentence furigana field', () => {
  const integration = createIntegration({
    fields: { word: 'Expression', sentence: 'Sentence' },
    behavior: { highlightWord: true },
  });

  const processed = internalsOf(integration).processSentenceFurigana(
    '<span class="term"><ruby>不思議<rt>ふしぎ</rt></ruby></span><span class="term">な</span><span class="term"><ruby>特技<rt>とくぎ</rt></ruby></span><span class="term">を</span>',
    { expression: '特技', sentence: '不思議な特技を' },
  );

  assert.equal(
    processed,
    '<span class="term"><ruby>不思議<rt>ふしぎ</rt></ruby></span><span class="term">な</span><b><span class="term"><ruby>特技<rt>とくぎ</rt></ruby></span></b><span class="term">を</span>',
  );
});

// --- Proxy, field grouping, notifications -----------------------------------

test('AnkiIntegration reports an occupied proxy address through its notification seam', async () => {
  const occupiedServer = http.createServer();
  occupiedServer.listen(0, '127.0.0.1');
  await once(occupiedServer, 'listening');
  const occupiedAddress = occupiedServer.address();
  assert.ok(occupiedAddress && typeof occupiedAddress === 'object');
  const stateDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-anki-proxy-collision-'));
  const overlayNotifications: TestOverlayNotificationPayload[] = [];
  const integration = createIntegration(
    {
      enabled: true,
      url: 'http://127.0.0.1:8765',
      proxy: {
        enabled: true,
        host: '127.0.0.1',
        port: occupiedAddress.port,
        upstreamUrl: 'http://127.0.0.1:8765',
      },
      behavior: { notificationType: 'overlay' },
      knownWords: { highlightEnabled: false },
      nPlusOne: { enabled: false },
    },
    {
      knownWordCacheStatePath: path.join(stateDir, 'known-words-cache.json'),
      overlay: (payload) => {
        overlayNotifications.push(payload);
      },
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
  const integration = createIntegration({
    isKiku: { enabled: true, fieldGrouping: 'manual' },
  });
  let groupingTriggered = 0;
  const internals = internalsOf(integration);
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
  const integration = createIntegration(
    { behavior: { notificationType: 'osd' } },
    { osd: (text) => osdMessages.push(text) },
  );

  await internalsOf(integration).showNotification(42, 'taberu', 'image failed');

  assert.deepEqual(osdMessages, ['x Updated card: taberu (image failed)']);
});

test('AnkiIntegration routes workflow status notifications through configured surfaces', async () => {
  const osdMessages: string[] = [];
  const desktopMessages: string[] = [];
  const overlayMessages: string[] = [];
  const integration = createIntegration(
    { behavior: { notificationType: 'both' } },
    {
      osd: (text) => osdMessages.push(text),
      desktop: (title, options) => desktopMessages.push(`${title}:${options.body ?? ''}`),
      overlay: (payload) =>
        overlayMessages.push(`${payload.title}:${payload.body ?? ''}:${payload.variant ?? ''}`),
    },
  );

  assert.equal(await integration.createSentenceCard('食べる', 0, 1), false);

  assert.deepEqual(osdMessages, []);
  assert.deepEqual(overlayMessages, ['SubMiner:No video loaded:info']);
  assert.deepEqual(desktopMessages, ['SubMiner:No video loaded']);
});

// --- YouTube media cache ----------------------------------------------------

/**
 * Installs a fake AnkiConnect client whose notes carry empty values for `noteFieldNames`, plus a
 * media generator that succeeds unless overridden. Records every note update and stored file.
 */
function installFakeAnki(
  integration: AnkiIntegration,
  options: {
    noteFieldNames: string[];
    mediaGenerator?: IntegrationInternals['mediaGenerator'];
  },
) {
  const updatedNotes: Array<{ noteId: number; fields: Record<string, string> }> = [];
  const storedMedia: string[] = [];
  const internals = internalsOf(integration);
  internals.client = {
    notesInfo: async (noteIds) =>
      noteIds.map((noteId) => ({
        noteId,
        fields: Object.fromEntries(options.noteFieldNames.map((name) => [name, { value: '' }])),
      })),
    updateNoteFields: async (noteId, fields) => {
      updatedNotes.push({ noteId, fields });
    },
    storeMediaFile: async (filename) => {
      storedMedia.push(filename);
    },
  };
  internals.mediaGenerator = {
    generateAudio: async () => Buffer.from('audio'),
    generateScreenshot: async () => Buffer.from('image'),
    ...options.mediaGenerator,
  };
  return { updatedNotes, storedMedia };
}

function youtubeJob(overrides: Partial<PendingYoutubeJob> & { noteId: number }): PendingYoutubeJob {
  return {
    sourceUrl: 'https://www.youtube.com/watch?v=abc123',
    startTime: 10,
    endTime: 12,
    label: String(overrides.noteId),
    audioFieldName: 'SentenceAudio',
    imageFieldName: 'Picture',
    generateAudio: true,
    generateImage: true,
    ...overrides,
  };
}

test('AnkiIntegration applies ready YouTube cache media to every queued note id', async () => {
  const osdMessages: string[] = [];
  const mediaInputs: string[] = [];
  const integration = createIntegration(
    {
      fields: { audio: 'ExpressionAudio', image: 'Picture' },
      media: { imageFormat: 'jpg' },
      behavior: { notificationType: 'osd' },
    },
    { osd: (text) => osdMessages.push(text) },
  );
  const { updatedNotes, storedMedia } = installFakeAnki(integration, {
    noteFieldNames: ['ExpressionAudio', 'Picture'],
    mediaGenerator: {
      generateAudio: async (mediaPath, _start, _end, _padding, audioStreamIndex, _norm, scale) => {
        mediaInputs.push(
          `audio:${describeMediaInputForTest(mediaPath)}:${audioStreamIndex ?? 'auto'}:${scale}`,
        );
        return Buffer.from('audio');
      },
      generateScreenshot: async (mediaPath) => {
        mediaInputs.push(`image:${describeMediaInputForTest(mediaPath)}`);
        return Buffer.from('image');
      },
    },
  });
  const internals = internalsOf(integration);
  internals.queuePendingYoutubeMediaUpdate(
    youtubeJob({ noteId: 101, label: 'first', audioStreamIndex: 22, volumeScale: 0.25 }),
  );
  internals.queuePendingYoutubeMediaUpdate(
    youtubeJob({
      noteId: 202,
      sourceUrl: 'https://youtu.be/abc123',
      startTime: 20,
      endTime: 22,
      label: 'second',
      audioStreamIndex: 23,
      volumeScale: 0.8,
    }),
  );

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
  for (const update of updatedNotes) {
    assert.match(update.fields.SentenceAudio ?? '', /^\[sound:audio_/);
    assert.match(update.fields.Picture ?? '', /^<img src="image_/);
  }
  assert.equal(storedMedia.length, 4);
  assert.ok(
    osdMessages.some((message) =>
      message.includes('YouTube media cache ready. Adding media to 2 queued cards.'),
    ),
  );
});

test('AnkiIntegration reports partial queued YouTube media updates separately from failures', async () => {
  const osdMessages: string[] = [];
  const notifications: Array<{ noteId: number; label: string | number; suffix?: string }> = [];
  const integration = createIntegration(
    {
      fields: { image: 'Picture' },
      media: { imageFormat: 'jpg' },
      behavior: { notificationType: 'osd' },
    },
    { osd: (text) => osdMessages.push(text) },
  );
  const { updatedNotes } = installFakeAnki(integration, {
    noteFieldNames: ['SentenceAudio', 'Picture'],
    mediaGenerator: {
      generateAudio: async () => {
        throw new Error('audio stream not found');
      },
    },
  });
  const internals = internalsOf(integration);
  internals.showNotification = async (noteId, label, suffix) => {
    notifications.push({ noteId, label, suffix });
  };
  internals.queuePendingYoutubeMediaUpdate(
    youtubeJob({
      noteId: 303,
      sourceUrl: 'https://www.youtube.com/watch?v=partial',
      label: 'partial',
    }),
  );

  await integration.handleYoutubeMediaCacheReady('https://youtu.be/partial', '/tmp/media.mkv');

  assert.equal(updatedNotes.length, 1);
  assert.match(updatedNotes[0]?.fields.Picture ?? '', /^<img src="image_/);
  assert.equal(updatedNotes[0]?.fields.SentenceAudio, undefined);
  assert.deepEqual(notifications, [{ noteId: 303, label: 'partial', suffix: 'audio failed' }]);
  assert.ok(
    osdMessages.some((message) =>
      message.includes('Queued YouTube media finished with 0 updated, 1 partial, and 0 failed.'),
    ),
  );
});

test('AnkiIntegration queues YouTube media updates against recovered source URLs', async () => {
  const audioVolumeScales: Array<number | undefined> = [];
  let mpvVolume = 30;
  const integration = createIntegration(
    { fields: { image: 'Picture' }, media: { imageFormat: 'jpg' } },
    {
      mpvClient: {
        currentVideoPath:
          'https://rr1---sn.example.googlevideo.com/videoplayback?expire=1777777777',
        currentSubStart: 10,
        currentSubEnd: 12,
        currentTimePos: 11,
        requestProperty: async (name: string) => {
          assert.equal(name, 'volume');
          return mpvVolume;
        },
      },
      osd: () => undefined,
      getCachedMediaPath: async () => null,
      shouldRequireRemoteMediaCache: () => true,
      getYoutubeMediaSourceUrl: () => 'https://www.youtube.com/watch?v=abc123',
    },
  );
  const { updatedNotes, storedMedia } = installFakeAnki(integration, {
    noteFieldNames: ['SentenceAudio', 'Picture'],
    mediaGenerator: {
      generateAudio: async (_path, _start, _end, _padding, _stream, _norm, volumeScale) => {
        audioVolumeScales.push(volumeScale);
        return Buffer.from('audio');
      },
    },
  });
  const internals = internalsOf(integration);
  internals.showNotification = async () => undefined;

  const queued = await internals.queuePendingYoutubeMediaUpdateForNote({
    noteId: 404,
    noteInfo: {
      noteId: 404,
      fields: { ExpressionAudio: { value: '' }, Picture: { value: '' } },
    },
    label: 'resolved source',
  });
  // The volume is captured when the job is queued, not when the cache becomes ready.
  mpvVolume = 90;
  await integration.handleYoutubeMediaCacheReady('https://youtu.be/abc123', '/tmp/media.mkv');

  assert.equal(queued, true);
  assert.equal(updatedNotes.length, 1);
  assert.equal(updatedNotes[0]?.noteId, 404);
  assert.match(updatedNotes[0]?.fields.ExpressionAudio ?? '', /^\[sound:audio_/);
  assert.equal(updatedNotes[0]?.fields.SentenceAudio, undefined);
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
  const integration = createIntegration(
    { media: { audioPadding: 0, normalizeAudio: false } },
    {
      mpvClient: {
        currentVideoPath: 'https://www.youtube.com/watch?v=abc123',
        currentAudioStreamIndex: 0,
        currentSubStart: 10,
        currentSubEnd: 12,
        currentTimePos: 11,
        requestProperty: async (name: string) => {
          requestedProperties.push(name);
          return 55;
        },
      },
      osd: () => undefined,
      getCachedMediaPath: async () => '/tmp/subminer-youtube-media-cache/media.mkv',
      shouldRequireRemoteMediaCache: () => true,
    },
  );
  internalsOf(integration).mediaGenerator = {
    generateAudio: async (
      mediaPath,
      _start,
      _end,
      _padding,
      audioStreamIndex,
      normalizeAudio,
      volumeScale,
    ) => {
      audioCalls.push({
        path: typeof mediaPath === 'string' ? mediaPath : mediaPath.path,
        audioStreamIndex,
        normalizeAudio,
        volumeScale,
      });
      return Buffer.from('audio');
    },
  };

  await internalsOf(integration).generateAudio();

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

const NO_QUEUED_CACHE_READY_CASES: Array<{
  name: string;
  options?: { notifyNoQueued: boolean };
  expectedBodies: string[];
}> = [
  {
    name: 'announces ready YouTube cache when no queued notes exist',
    expectedBodies: ['YouTube media cache ready.'],
  },
  {
    name: 'can let caller own no-queued YouTube cache ready notification',
    options: { notifyNoQueued: false },
    expectedBodies: [],
  },
];

for (const c of NO_QUEUED_CACHE_READY_CASES) {
  test(`AnkiIntegration ${c.name}`, async () => {
    const overlayNotifications: TestOverlayNotificationPayload[] = [];
    const integration = createIntegration(
      { behavior: { notificationType: 'overlay' } },
      { overlay: (payload) => overlayNotifications.push(payload) },
    );

    await integration.handleYoutubeMediaCacheReady(
      'https://youtu.be/abc123',
      '/tmp/media.mkv',
      c.options,
    );

    assert.deepEqual(
      overlayNotifications.map((notification) => [notification.title, notification.body]),
      c.expectedBodies.map((body) => ['SubMiner', body]),
    );
  });
}

// --- Overlay / desktop notification surfaces --------------------------------

/**
 * Integration in 'both' notification mode with a fake media generator whose temp-icon write is
 * `writeNotificationIconToFile`. Returns the surfaces' captured output.
 */
function createNotificationImageFixture(
  writeNotificationIconToFile: (iconBuffer: Buffer, noteId: number) => string,
) {
  const desktopNotifications: Array<{ title: string; body?: string; icon?: string }> = [];
  const overlayNotifications: TestOverlayNotificationPayload[] = [];
  const generatedFrom: Array<{ videoPath: string; timestamp: number }> = [];
  const cleanupPaths: string[] = [];
  const integration = createIntegration(
    { behavior: { notificationType: 'both' } },
    {
      mpvClient: { currentVideoPath: '/tmp/show.mkv', currentTimePos: 123.45 },
      desktop: (title, options) => {
        desktopNotifications.push({ title, body: options.body, icon: options.icon });
      },
      overlay: (payload) => overlayNotifications.push(payload),
    },
  );
  internalsOf(integration).mediaGenerator = {
    generateNotificationIcon: async (videoPath, timestamp) => {
      generatedFrom.push({ videoPath, timestamp });
      return Buffer.from('png');
    },
    writeNotificationIconToFile,
    scheduleNotificationIconCleanup: (filePath) => {
      cleanupPaths.push(filePath);
    },
  };
  return { integration, desktopNotifications, overlayNotifications, generatedFrom, cleanupPaths };
}

const PNG_DATA_URL = `data:image/png;base64,${Buffer.from('png').toString('base64')}`;

test('AnkiIntegration embeds generated notification image on overlay mined-card notifications', async () => {
  const notificationIconPath = path.join(os.tmpdir(), 'subminer-notification-icon.png');
  const fixture = createNotificationImageFixture((iconBuffer, noteId) => {
    assert.equal(iconBuffer.toString(), 'png');
    assert.equal(noteId, 42);
    return notificationIconPath;
  });

  await internalsOf(fixture.integration).showNotification(42, '食べる');

  assert.deepEqual(fixture.generatedFrom, [{ videoPath: '/tmp/show.mkv', timestamp: 123.45 }]);
  assert.equal(fixture.overlayNotifications.length, 1);
  assert.equal(fixture.overlayNotifications[0]?.title, 'Anki Card Updated');
  assert.equal(fixture.overlayNotifications[0]?.body, 'Updated card: 食べる');
  assert.equal(fixture.overlayNotifications[0]?.image, PNG_DATA_URL);
  assert.deepEqual(fixture.overlayNotifications[0]?.actions, [
    { id: 'open-anki-card', label: 'Open in Anki', noteId: 42 },
  ]);
  assert.deepEqual(fixture.desktopNotifications, [
    { title: 'Anki Card Updated', body: 'Updated card: 食べる', icon: notificationIconPath },
  ]);
  assert.deepEqual(fixture.cleanupPaths, [notificationIconPath]);
});

test('AnkiIntegration keeps overlay notification image when temp icon write fails', async () => {
  const fixture = createNotificationImageFixture(() => {
    throw new Error('disk full');
  });

  await internalsOf(fixture.integration).showNotification(42, '食べる');

  assert.equal(fixture.overlayNotifications[0]?.image, PNG_DATA_URL);
  assert.deepEqual(fixture.desktopNotifications, [
    { title: 'Anki Card Updated', body: 'Updated card: 食べる', icon: undefined },
  ]);
  assert.deepEqual(fixture.cleanupPaths, []);
});

test('AnkiIntegration keeps overlay card-update progress visible until the terminal notification', async () => {
  const overlayNotifications: TestOverlayNotificationPayload[] = [];
  const integration = createIntegration(
    { behavior: { notificationType: 'overlay' } },
    { overlay: (payload) => overlayNotifications.push(payload) },
  );
  const internals = internalsOf(integration);

  internals.beginUpdateProgress('Updating card');
  await internals.showNotification(42, '食べる');

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
  const integration = createIntegration(
    { behavior: { notificationType: 'overlay' } },
    {
      overlay: (payload) => overlayNotifications.push(payload),
      dismissOverlay: (id) => dismissedIds.push(id),
    },
  );
  const internals = internalsOf(integration);

  internals.beginUpdateProgress('Updating card');
  internals.endUpdateProgress();

  assert.equal(overlayNotifications[0]?.persistent, true);
  assert.deepEqual(dismissedIds, ['anki-update-progress']);
});

test('AnkiIntegration dismisses overlay update progress after notifications switch to OSD', () => {
  const behavior: NonNullable<AnkiConnectConfig['behavior']> = { notificationType: 'overlay' };
  const dismissedIds: string[] = [];
  const integration = createIntegration(
    { behavior },
    { overlay: () => {}, dismissOverlay: (id) => dismissedIds.push(id) },
  );
  const internals = internalsOf(integration);

  internals.beginUpdateProgress('Updating card');
  behavior.notificationType = 'osd';
  internals.endUpdateProgress();

  assert.deepEqual(dismissedIds, ['anki-update-progress']);
});

// --- Misc-info metadata -----------------------------------------------------

const MISC_INFO_CASES = [
  {
    name: 'avoids leaking Jellyfin api_key query params',
    pattern: '[SubMiner] %f (%t)',
    mpv: {
      currentVideoPath:
        'stream?static=true&api_key=secret-token&MediaSourceId=a762ab23d26d4347e3cacdb83aaae405&AudioStreamIndex=3',
      currentMediaTitle: '[Jellyfin/direct] Bocchi the Rock! - S01E02',
    },
    fallbackFilename: 'audio_123.mp3',
    expected: '[SubMiner] [Jellyfin/direct] Bocchi the Rock! - S01E02 (00:07:06)',
    leaked: 'api_key=',
  },
  {
    name: 'treats ApiKey stream paths like legacy api_key ones',
    pattern: '[SubMiner] %f (%t)',
    mpv: {
      currentVideoPath: 'stream?static=true&ApiKey=secret-token&MediaSourceId=ms-1',
      currentMediaTitle: '[Jellyfin/direct] Bocchi the Rock! - S01E02',
    },
    fallbackFilename: 'audio_123.mp3',
    expected: '[SubMiner] [Jellyfin/direct] Bocchi the Rock! - S01E02 (00:07:06)',
    leaked: 'ApiKey=',
  },
  {
    name: 'rejects a credential-bearing media title before metadata arrives',
    pattern: '[SubMiner] %f | %F (%t)',
    mpv: {
      currentVideoPath: 'https://jellyfin.example/Videos/item/stream?api_key=test-secret',
      currentMediaTitle: 'stream?static=true&api_key=test-secret',
    },
    fallbackFilename: 'stream?api_key=test-secret',
    expected: '[SubMiner] Unknown media | Unknown media (00:07:06)',
    leaked: 'test-secret',
  },
];

for (const c of MISC_INFO_CASES) {
  test(`AnkiIntegration.formatMiscInfoPattern ${c.name}`, () => {
    const integration = createIntegration(
      { metadata: { pattern: c.pattern } },
      {
        mpvClient: {
          currentSubText: '',
          currentTimePos: 426,
          currentSubStart: 426,
          currentSubEnd: 428,
          send: () => true,
          ...c.mpv,
        },
      },
    );

    const result = internalsOf(integration).formatMiscInfoPattern(c.fallbackFilename, 426);

    assert.equal(result, c.expected);
    assert.equal(result.includes(c.leaked), false);
  });
}

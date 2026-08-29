import assert from 'node:assert/strict';
import test from 'node:test';

import { CardCreationService } from './card-creation';
import { toMpvEdlValue } from './mpv-edl-test-utils';
import type { MediaInput } from '../media-generator';
import type { AnkiConnectConfig, CardKind } from '../types/anki';
import { applyCardKindFlagFields } from './card-kinds';

type CardCreationDeps = ConstructorParameters<typeof CardCreationService>[0];

function setCardTypeFields(
  updatedFields: Record<string, string>,
  availableFieldNames: string[],
  cardKind: CardKind,
): void {
  applyCardKindFlagFields(
    updatedFields,
    cardKind,
    (preferredName) =>
      availableFieldNames.find((name) => name.toLowerCase() === preferredName.toLowerCase()) ??
      null,
  );
}

function createManualUpdateService(overrides: Partial<CardCreationDeps> = {}): {
  service: CardCreationService;
  updatedFields: Record<string, string>[];
  mergeCalls: Array<{ existing: string; newValue: string; overwrite: boolean }>;
  storedMedia: string[];
} {
  const updatedFields: Record<string, string>[] = [];
  const mergeCalls: Array<{ existing: string; newValue: string; overwrite: boolean }> = [];
  const storedMedia: string[] = [];

  const deps: CardCreationDeps = {
    getConfig: () =>
      ({
        deck: 'Mining',
        fields: {
          word: 'Expression',
          sentence: 'Sentence',
          audio: 'ExpressionAudio',
        },
        media: {
          generateAudio: true,
          generateImage: false,
          maxMediaDuration: 30,
        },
        behavior: {
          overwriteAudio: false,
          overwriteImage: false,
        },
        ai: false,
      }) as AnkiConnectConfig,
    getAiConfig: () => ({}),
    getTimingTracker: () =>
      ({
        findTiming: (text: string) => (text === '字幕' ? { startTime: 12, endTime: 14 } : null),
      }) as never,
    getMpvClient: () =>
      ({
        currentVideoPath: '/video.mp4',
        currentAudioStreamIndex: 0,
      }) as never,
    client: {
      addNote: async () => 0,
      addTags: async () => undefined,
      notesInfo: async () => [
        {
          noteId: 42,
          fields: {
            Expression: { value: '単語' },
            Sentence: { value: '' },
            ExpressionAudio: { value: '[sound:auto-expression.mp3]' },
            SentenceAudio: { value: '[sound:auto-sentence.mp3]' },
          },
        },
      ],
      updateNoteFields: async (_noteId, fields) => {
        updatedFields.push(fields);
      },
      storeMediaFile: async (filename) => {
        storedMedia.push(filename);
      },
      findNotes: async () => [42],
      retrieveMediaFile: async () => '',
    },
    mediaGenerator: {
      generateAudio: async () => Buffer.from('audio'),
      generateScreenshot: async () => null,
      generateAnimatedImage: async () => null,
    },
    showOsdNotification: () => undefined,
    showUpdateResult: () => undefined,
    showStatusNotification: () => undefined,
    showNotification: async () => undefined,
    beginUpdateProgress: () => undefined,
    endUpdateProgress: () => undefined,
    withUpdateProgress: async (_message, action) => action(),
    resolveConfiguredFieldName: (noteInfo, ...preferredNames) => {
      for (const preferredName of preferredNames) {
        if (preferredName && preferredName in noteInfo.fields) return preferredName;
      }
      return null;
    },
    resolveNoteFieldName: (noteInfo, preferredName) =>
      preferredName && preferredName in noteInfo.fields ? preferredName : null,
    getAnimatedImageLeadInSeconds: async () => 0,
    extractFields: (fields) =>
      Object.fromEntries(
        Object.entries(fields).map(([name, field]) => [name.toLowerCase(), field.value]),
      ),
    processSentence: (sentence) => sentence,
    setCardTypeFields: () => undefined,
    mergeFieldValue: (existing, newValue, overwrite) => {
      mergeCalls.push({ existing, newValue, overwrite });
      return overwrite || !existing.trim() ? newValue : existing;
    },
    formatMiscInfoPattern: () => '',
    getEffectiveSentenceCardConfig: () => ({
      model: 'Sentence',
      sentenceField: 'Sentence',
      audioField: 'SentenceAudio',
      lapisEnabled: false,
      kikuEnabled: false,
      kikuFieldGrouping: 'disabled',
      kikuDeleteDuplicateInAuto: false,
    }),
    getFallbackDurationSeconds: () => 10,
    appendKnownWordsFromNoteInfo: () => undefined,
    isUpdateInProgress: () => false,
    setUpdateInProgress: () => undefined,
    trackLastAddedNoteId: () => undefined,
    ...overrides,
  };

  return {
    service: new CardCreationService(deps),
    updatedFields,
    mergeCalls,
    storedMedia,
  };
}

test('manual clipboard subtitle update replaces audio in the configured field', async () => {
  const { service, updatedFields, mergeCalls, storedMedia } = createManualUpdateService();

  await service.updateLastAddedFromClipboard('字幕');

  assert.equal(updatedFields.length, 1);
  assert.equal(storedMedia.length, 1);
  const audioValue = `[sound:${storedMedia[0]}]`;
  assert.equal(updatedFields[0]?.ExpressionAudio, audioValue);
  assert.equal('SentenceAudio' in updatedFields[0]!, false);
  assert.deepEqual(
    mergeCalls.map((call) => call.overwrite),
    [true],
  );
});

test('manual clipboard word-card update uses configured fields with Lapis and Kiku enabled', async () => {
  const { service, updatedFields } = createManualUpdateService({
    getConfig: () =>
      ({
        deck: 'Mining',
        fields: {
          word: 'Expression',
          sentence: 'Context',
          audio: 'ContextAudio',
        },
        media: {
          generateAudio: true,
          generateImage: false,
          maxMediaDuration: 30,
        },
        behavior: {
          overwriteAudio: false,
          overwriteImage: false,
        },
        ai: false,
      }) as AnkiConnectConfig,
    client: {
      addNote: async () => 0,
      addTags: async () => undefined,
      notesInfo: async () => [
        {
          noteId: 42,
          fields: {
            Expression: { value: '単語' },
            Sentence: { value: '' },
            SentenceAudio: { value: '' },
            Context: { value: '' },
            ContextAudio: { value: '' },
          },
        },
      ],
      updateNoteFields: async (_noteId, fields) => {
        updatedFields.push(fields);
      },
      storeMediaFile: async () => undefined,
      findNotes: async () => [42],
      retrieveMediaFile: async () => '',
    },
    getEffectiveSentenceCardConfig: () => ({
      model: 'Sentence',
      sentenceField: 'Sentence',
      audioField: 'SentenceAudio',
      lapisEnabled: true,
      kikuEnabled: true,
      kikuFieldGrouping: 'disabled',
      kikuDeleteDuplicateInAuto: false,
    }),
  });

  await service.updateLastAddedFromClipboard('字幕');

  assert.equal(updatedFields.length, 1);
  assert.match(updatedFields[0]?.ContextAudio ?? '', /^\[sound:audio_\d+\.mp3\]$/);
  assert.deepEqual(Object.keys(updatedFields[0] ?? {}).sort(), ['Context', 'ContextAudio']);
  assert.equal(updatedFields[0]?.Context, '字幕');
});

test('audio-card action keeps Lapis and Kiku sentence fields', async () => {
  const { service, updatedFields } = createManualUpdateService({
    getConfig: () =>
      ({
        deck: 'Mining',
        fields: {
          word: 'Expression',
          sentence: 'Context',
          audio: 'ContextAudio',
        },
        media: {
          generateAudio: true,
          generateImage: false,
          maxMediaDuration: 30,
        },
        behavior: {},
        ai: false,
      }) as AnkiConnectConfig,
    getMpvClient: () =>
      ({
        currentVideoPath: '/video.mp4',
        currentAudioStreamIndex: 0,
        currentSubText: '字幕',
        currentSubStart: 12,
        currentSubEnd: 14,
      }) as never,
    client: {
      addNote: async () => 0,
      addTags: async () => undefined,
      notesInfo: async () => [
        {
          noteId: 42,
          fields: {
            Expression: { value: '単語' },
            Sentence: { value: '' },
            SentenceAudio: { value: '' },
            Context: { value: '' },
            ContextAudio: { value: '' },
          },
        },
      ],
      updateNoteFields: async (_noteId, fields) => {
        updatedFields.push(fields);
      },
      storeMediaFile: async () => undefined,
      findNotes: async () => [42],
      retrieveMediaFile: async () => '',
    },
    getEffectiveSentenceCardConfig: () => ({
      model: 'Sentence',
      sentenceField: 'Sentence',
      audioField: 'SentenceAudio',
      lapisEnabled: true,
      kikuEnabled: true,
      kikuFieldGrouping: 'disabled',
      kikuDeleteDuplicateInAuto: false,
    }),
  });

  await service.markLastCardAsAudioCard();

  assert.equal(updatedFields.length, 1);
  assert.equal(updatedFields[0]?.Sentence, '字幕');
  assert.match(updatedFields[0]?.SentenceAudio ?? '', /^\[sound:audio_\d+\.mp3\]$/);
  assert.equal('Context' in (updatedFields[0] ?? {}), false);
  assert.equal('ContextAudio' in (updatedFields[0] ?? {}), false);
});

test('manual clipboard subtitle update marks Kiku word cards as word-and-sentence cards when enabled', async () => {
  const { service, updatedFields } = createManualUpdateService({
    getConfig: () =>
      ({
        deck: 'Mining',
        fields: {
          word: 'Expression',
          sentence: 'Sentence',
          audio: 'ExpressionAudio',
        },
        media: {
          generateAudio: false,
          generateImage: false,
          maxMediaDuration: 30,
        },
        behavior: {
          overwriteAudio: false,
          overwriteImage: false,
        },
        ai: false,
      }) as AnkiConnectConfig,
    client: {
      addNote: async () => 0,
      addTags: async () => undefined,
      notesInfo: async () => [
        {
          noteId: 42,
          fields: {
            Expression: { value: '単語' },
            Sentence: { value: '' },
            IsWordAndSentenceCard: { value: '' },
            IsSentenceCard: { value: '' },
            IsAudioCard: { value: '' },
          },
        },
      ],
      updateNoteFields: async (_noteId, fields) => {
        updatedFields.push(fields);
      },
      storeMediaFile: async () => undefined,
      findNotes: async () => [42],
      retrieveMediaFile: async () => '',
    },
    getEffectiveSentenceCardConfig: () => ({
      model: 'Sentence',
      sentenceField: 'Sentence',
      audioField: 'SentenceAudio',
      lapisEnabled: false,
      kikuEnabled: true,
      kikuFieldGrouping: 'disabled',
      kikuDeleteDuplicateInAuto: false,
    }),
    setCardTypeFields,
  });

  await service.updateLastAddedFromClipboard('字幕');

  assert.equal(updatedFields.length, 1);
  assert.deepEqual(updatedFields[0], {
    Sentence: '字幕',
    IsWordAndSentenceCard: 'x',
    IsSentenceCard: '',
    IsAudioCard: '',
  });
});

test('manual clipboard subtitle update uses configured audio when SentenceAudio is missing', async () => {
  const { service, updatedFields, mergeCalls, storedMedia } = createManualUpdateService({
    client: {
      addNote: async () => 0,
      addTags: async () => undefined,
      notesInfo: async () => [
        {
          noteId: 42,
          fields: {
            Expression: { value: '単語' },
            Sentence: { value: '' },
            ExpressionAudio: { value: '[sound:auto-expression.mp3]' },
          },
        },
      ],
      updateNoteFields: async (_noteId, fields) => {
        updatedFields.push(fields);
      },
      storeMediaFile: async (filename) => {
        storedMedia.push(filename);
      },
      findNotes: async () => [42],
      retrieveMediaFile: async () => '',
    },
  });

  await service.updateLastAddedFromClipboard('字幕');

  assert.equal(storedMedia.length, 1);
  assert.equal(updatedFields.length, 1);
  assert.match(updatedFields[0]?.ExpressionAudio ?? '', /^\[sound:audio_\d+\.mp3\]$/);
  assert.equal(updatedFields[0]?.Sentence, '字幕');
  assert.equal(mergeCalls.length, 1);
});

test('manual clipboard subtitle update uses resolved mpv stream URLs for remote media', async () => {
  const audioPaths: string[] = [];
  const imagePaths: string[] = [];
  const recordMediaPath = (mediaInput: MediaInput): string =>
    typeof mediaInput === 'string' ? mediaInput : mediaInput.path;
  const audioUrl = 'https://audio.example/videoplayback?mime=audio%2Fwebm';
  const videoUrl = 'https://video.example/videoplayback?mime=video%2Fmp4';
  const edlSource = [
    `edl://!new_stream;!no_clip;!no_chapters;${toMpvEdlValue(audioUrl)}`,
    `!new_stream;!no_clip;!no_chapters;${toMpvEdlValue(videoUrl)}`,
    '!global_tags,title=test',
  ].join(';');

  const { service, updatedFields, storedMedia } = createManualUpdateService({
    getConfig: () =>
      ({
        deck: 'Mining',
        fields: {
          word: 'Expression',
          sentence: 'Sentence',
          audio: 'ExpressionAudio',
          image: 'Picture',
        },
        media: {
          generateAudio: true,
          generateImage: true,
          imageFormat: 'jpg',
          maxMediaDuration: 30,
        },
        behavior: {
          overwriteAudio: false,
          overwriteImage: false,
        },
        ai: false,
      }) as AnkiConnectConfig,
    getTimingTracker: () =>
      ({
        findTiming: (text: string) => {
          if (text === '一行目') return { startTime: 10, endTime: 12 };
          if (text === '二行目') return { startTime: 12.5, endTime: 14 };
          return null;
        },
      }) as never,
    getMpvClient: () =>
      ({
        currentVideoPath: 'https://www.youtube.com/watch?v=abc123',
        currentTimePos: 13,
        currentAudioStreamIndex: 0,
        requestProperty: async (name: string) => {
          assert.equal(name, 'stream-open-filename');
          return edlSource;
        },
      }) as never,
    client: {
      addNote: async () => 0,
      addTags: async () => undefined,
      notesInfo: async () => [
        {
          noteId: 42,
          fields: {
            Expression: { value: '単語' },
            Sentence: { value: '' },
            ExpressionAudio: { value: '[sound:auto-expression.mp3]' },
            SentenceAudio: { value: '[sound:auto-sentence.mp3]' },
            Picture: { value: '' },
          },
        },
      ],
      updateNoteFields: async (_noteId, fields) => {
        updatedFields.push(fields);
      },
      storeMediaFile: async (filename) => {
        storedMedia.push(filename);
      },
      findNotes: async () => [42],
      retrieveMediaFile: async () => '',
    },
    mediaGenerator: {
      generateAudio: async (path) => {
        audioPaths.push(recordMediaPath(path));
        return Buffer.from('audio');
      },
      generateScreenshot: async (path) => {
        imagePaths.push(recordMediaPath(path));
        return Buffer.from('image');
      },
      generateAnimatedImage: async () => null,
    },
  });

  await service.updateLastAddedFromClipboard('一行目\n\n二行目');

  assert.deepEqual(audioPaths, [audioUrl]);
  assert.deepEqual(imagePaths, [videoUrl]);
  assert.equal(storedMedia.length, 2);
  assert.equal(updatedFields.length, 1);
  assert.equal(updatedFields[0]?.Sentence, '一行目 二行目');
  assert.match(updatedFields[0]?.Picture ?? '', /^<img src="image_\d+\.jpg">$/);
});

test('createSentenceCard relies on Anki progress notification without standalone status toast', async () => {
  const statusMessages: string[] = [];
  const progressMessages: string[] = [];
  const { service } = createManualUpdateService({
    showOsdNotification: (message) => {
      statusMessages.push(message);
    },
    withUpdateProgress: async (message, action) => {
      progressMessages.push(message);
      return await action();
    },
    mediaGenerator: {
      generateAudio: async () => null,
      generateScreenshot: async () => null,
      generateAnimatedImage: async () => null,
    },
  });

  const created = await service.createSentenceCard('テスト', 0, 1);

  assert.equal(created, true);
  assert.deepEqual(progressMessages, ['Creating sentence card']);
  assert.deepEqual(statusMessages, []);
});

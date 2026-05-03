import assert from 'node:assert/strict';
import test from 'node:test';

import { CardCreationService } from './card-creation';
import type { AnkiConnectConfig } from '../types/anki';

type CardCreationDeps = ConstructorParameters<typeof CardCreationService>[0];

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

test('manual clipboard subtitle update replaces sentence audio without touching expression audio', async () => {
  const { service, updatedFields, mergeCalls, storedMedia } = createManualUpdateService();

  await service.updateLastAddedFromClipboard('字幕');

  assert.equal(updatedFields.length, 1);
  assert.equal(storedMedia.length, 1);
  const audioValue = `[sound:${storedMedia[0]}]`;
  assert.equal(updatedFields[0]?.SentenceAudio, audioValue);
  assert.equal('ExpressionAudio' in updatedFields[0]!, false);
  assert.deepEqual(
    mergeCalls.map((call) => call.overwrite),
    [true],
  );
});

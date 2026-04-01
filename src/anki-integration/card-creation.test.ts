import assert from 'node:assert/strict';
import test from 'node:test';

import { CardCreationService } from './card-creation';
import type { AnkiConnectConfig } from '../types/anki';

test('CardCreationService counts locally created sentence cards', async () => {
  const minedCards: Array<{ count: number; noteIds?: number[] }> = [];
  const service = new CardCreationService({
    getConfig: () =>
      ({
        deck: 'Mining',
        fields: {
          sentence: 'Sentence',
          audio: 'SentenceAudio',
        },
        media: {
          generateAudio: false,
          generateImage: false,
        },
        behavior: {},
        ai: false,
      }) as AnkiConnectConfig,
    getAiConfig: () => ({}),
    getTimingTracker: () => ({}) as never,
    getMpvClient: () =>
      ({
        currentVideoPath: '/video.mp4',
        currentSubText: '字幕',
        currentSubStart: 1,
        currentSubEnd: 2,
        currentTimePos: 1.5,
        currentAudioStreamIndex: 0,
      }) as never,
    client: {
      addNote: async () => 42,
      addTags: async () => undefined,
      notesInfo: async () => [],
      updateNoteFields: async () => undefined,
      storeMediaFile: async () => undefined,
      findNotes: async () => [],
      retrieveMediaFile: async () => '',
    },
    mediaGenerator: {
      generateAudio: async () => null,
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
    resolveConfiguredFieldName: () => null,
    resolveNoteFieldName: () => null,
    getAnimatedImageLeadInSeconds: async () => 0,
    extractFields: () => ({}),
    processSentence: (sentence) => sentence,
    setCardTypeFields: () => undefined,
    mergeFieldValue: (_existing, newValue) => newValue,
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
    recordCardsMinedCallback: (count, noteIds) => {
      minedCards.push({ count, noteIds });
    },
  });

  const created = await service.createSentenceCard('テスト', 0, 1);

  assert.equal(created, true);
  assert.deepEqual(minedCards, [{ count: 1, noteIds: [42] }]);
});

test('CardCreationService keeps updating after trackLastAddedNoteId throws', async () => {
  const calls = {
    notesInfo: 0,
    updateNoteFields: 0,
  };
  const service = new CardCreationService({
    getConfig: () =>
      ({
        deck: 'Mining',
        fields: {
          sentence: 'Sentence',
          audio: 'SentenceAudio',
        },
        media: {
          generateAudio: false,
          generateImage: false,
        },
        behavior: {},
        ai: false,
      }) as AnkiConnectConfig,
    getAiConfig: () => ({}),
    getTimingTracker: () => ({}) as never,
    getMpvClient: () =>
      ({
        currentVideoPath: '/video.mp4',
        currentSubText: '字幕',
        currentSubStart: 1,
        currentSubEnd: 2,
        currentTimePos: 1.5,
        currentAudioStreamIndex: 0,
      }) as never,
    client: {
      addNote: async () => 42,
      addTags: async () => undefined,
      notesInfo: async () => {
        calls.notesInfo += 1;
        return [
          {
            noteId: 42,
            fields: {
              Sentence: { value: 'existing' },
            },
          },
        ];
      },
      updateNoteFields: async () => {
        calls.updateNoteFields += 1;
      },
      storeMediaFile: async () => undefined,
      findNotes: async () => [],
      retrieveMediaFile: async () => '',
    },
    mediaGenerator: {
      generateAudio: async () => null,
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
    resolveConfiguredFieldName: () => null,
    resolveNoteFieldName: () => null,
    getAnimatedImageLeadInSeconds: async () => 0,
    extractFields: () => ({}),
    processSentence: (sentence) => sentence,
    setCardTypeFields: (updatedFields) => {
      updatedFields.CardType = 'sentence';
    },
    mergeFieldValue: (_existing, newValue) => newValue,
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
    trackLastAddedNoteId: () => {
      throw new Error('track failed');
    },
  });

  const created = await service.createSentenceCard('テスト', 0, 1);

  assert.equal(created, true);
  assert.equal(calls.notesInfo, 1);
  assert.equal(calls.updateNoteFields, 1);
});

test('CardCreationService keeps updating after recordCardsMinedCallback throws', async () => {
  const calls = {
    notesInfo: 0,
    updateNoteFields: 0,
  };
  const service = new CardCreationService({
    getConfig: () =>
      ({
        deck: 'Mining',
        fields: {
          sentence: 'Sentence',
          audio: 'SentenceAudio',
        },
        media: {
          generateAudio: false,
          generateImage: false,
        },
        behavior: {},
        ai: false,
      }) as AnkiConnectConfig,
    getAiConfig: () => ({}),
    getTimingTracker: () => ({}) as never,
    getMpvClient: () =>
      ({
        currentVideoPath: '/video.mp4',
        currentSubText: '字幕',
        currentSubStart: 1,
        currentSubEnd: 2,
        currentTimePos: 1.5,
        currentAudioStreamIndex: 0,
      }) as never,
    client: {
      addNote: async () => 42,
      addTags: async () => undefined,
      notesInfo: async () => {
        calls.notesInfo += 1;
        return [
          {
            noteId: 42,
            fields: {
              Sentence: { value: 'existing' },
            },
          },
        ];
      },
      updateNoteFields: async () => {
        calls.updateNoteFields += 1;
      },
      storeMediaFile: async () => undefined,
      findNotes: async () => [],
      retrieveMediaFile: async () => '',
    },
    mediaGenerator: {
      generateAudio: async () => null,
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
    resolveConfiguredFieldName: () => null,
    resolveNoteFieldName: () => null,
    getAnimatedImageLeadInSeconds: async () => 0,
    extractFields: () => ({}),
    processSentence: (sentence) => sentence,
    setCardTypeFields: (updatedFields) => {
      updatedFields.CardType = 'sentence';
    },
    mergeFieldValue: (_existing, newValue) => newValue,
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
    recordCardsMinedCallback: () => {
      throw new Error('record failed');
    },
  });

  const created = await service.createSentenceCard('テスト', 0, 1);

  assert.equal(created, true);
  assert.equal(calls.notesInfo, 1);
  assert.equal(calls.updateNoteFields, 1);
});

test('CardCreationService uses stream-open-filename for remote media generation', async () => {
  const audioPaths: string[] = [];
  const imagePaths: string[] = [];
  const edlSource = [
    'edl://!new_stream;!no_clip;!no_chapters;%70%https://audio.example/videoplayback?mime=audio%2Fwebm',
    '!new_stream;!no_clip;!no_chapters;%69%https://video.example/videoplayback?mime=video%2Fmp4',
    '!global_tags,title=test',
  ].join(';');

  const service = new CardCreationService({
    getConfig: () =>
      ({
        deck: 'Mining',
        fields: {
          sentence: 'Sentence',
          audio: 'SentenceAudio',
          image: 'Picture',
        },
        media: {
          generateAudio: true,
          generateImage: true,
          imageFormat: 'jpg',
        },
        behavior: {},
        ai: false,
      }) as AnkiConnectConfig,
    getAiConfig: () => ({}),
    getTimingTracker: () => ({}) as never,
    getMpvClient: () =>
      ({
        currentVideoPath: 'https://www.youtube.com/watch?v=abc123',
        currentSubText: '字幕',
        currentSubStart: 1,
        currentSubEnd: 2,
        currentTimePos: 1.5,
        currentAudioStreamIndex: 0,
        requestProperty: async (name: string) => {
          assert.equal(name, 'stream-open-filename');
          return edlSource;
        },
      }) as never,
    client: {
      addNote: async () => 42,
      addTags: async () => undefined,
      notesInfo: async () => [
        {
          noteId: 42,
          fields: {
            Sentence: { value: '' },
            SentenceAudio: { value: '' },
            Picture: { value: '' },
          },
        },
      ],
      updateNoteFields: async () => undefined,
      storeMediaFile: async () => undefined,
      findNotes: async () => [],
      retrieveMediaFile: async () => '',
    },
    mediaGenerator: {
      generateAudio: async (path) => {
        audioPaths.push(path);
        return Buffer.from('audio');
      },
      generateScreenshot: async (path) => {
        imagePaths.push(path);
        return Buffer.from('image');
      },
      generateAnimatedImage: async () => null,
    },
    showOsdNotification: () => undefined,
    showUpdateResult: () => undefined,
    showStatusNotification: () => undefined,
    showNotification: async () => undefined,
    beginUpdateProgress: () => undefined,
    endUpdateProgress: () => undefined,
    withUpdateProgress: async (_message, action) => action(),
    resolveConfiguredFieldName: (noteInfo, preferredName) => {
      if (!preferredName) return null;
      return Object.keys(noteInfo.fields).find((field) => field === preferredName) ?? null;
    },
    resolveNoteFieldName: (noteInfo, preferredName) => {
      if (!preferredName) return null;
      return Object.keys(noteInfo.fields).find((field) => field === preferredName) ?? null;
    },
    getAnimatedImageLeadInSeconds: async () => 0,
    extractFields: () => ({}),
    processSentence: (sentence) => sentence,
    setCardTypeFields: () => undefined,
    mergeFieldValue: (_existing, newValue) => newValue,
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
  });

  const created = await service.createSentenceCard('テスト', 0, 1);

  assert.equal(created, true);
  assert.deepEqual(audioPaths, ['https://audio.example/videoplayback?mime=audio%2Fwebm']);
  assert.deepEqual(imagePaths, ['https://video.example/videoplayback?mime=video%2Fmp4']);
});

test('CardCreationService tracks pre-add duplicate note ids for kiku sentence cards', async () => {
  const trackedDuplicates: Array<{ noteId: number; duplicateNoteIds: number[] }> = [];
  const duplicateLookupExpressions: string[] = [];

  const service = new CardCreationService({
    getConfig: () =>
      ({
        deck: 'Mining',
        fields: {
          word: 'Expression',
          sentence: 'Sentence',
          audio: 'SentenceAudio',
        },
        media: {
          generateAudio: false,
          generateImage: false,
        },
        behavior: {},
        ai: false,
      }) as AnkiConnectConfig,
    getAiConfig: () => ({}),
    getTimingTracker: () => ({}) as never,
    getMpvClient: () =>
      ({
        currentVideoPath: '/video.mp4',
        currentSubText: '字幕',
        currentSubStart: 1,
        currentSubEnd: 2,
        currentTimePos: 1.5,
        currentAudioStreamIndex: 0,
      }) as never,
    client: {
      addNote: async () => 42,
      addTags: async () => undefined,
      notesInfo: async () => [],
      updateNoteFields: async () => undefined,
      storeMediaFile: async () => undefined,
      findNotes: async () => [],
      retrieveMediaFile: async () => '',
    },
    mediaGenerator: {
      generateAudio: async () => null,
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
    resolveConfiguredFieldName: () => null,
    resolveNoteFieldName: () => null,
    getAnimatedImageLeadInSeconds: async () => 0,
    extractFields: () => ({}),
    processSentence: (sentence) => sentence,
    setCardTypeFields: () => undefined,
    mergeFieldValue: (_existing, newValue) => newValue,
    formatMiscInfoPattern: () => '',
    getEffectiveSentenceCardConfig: () => ({
      model: 'Sentence',
      sentenceField: 'Sentence',
      audioField: 'SentenceAudio',
      lapisEnabled: false,
      kikuEnabled: true,
      kikuFieldGrouping: 'manual',
      kikuDeleteDuplicateInAuto: false,
    }),
    getFallbackDurationSeconds: () => 10,
    appendKnownWordsFromNoteInfo: () => undefined,
    isUpdateInProgress: () => false,
    setUpdateInProgress: () => undefined,
    trackLastAddedNoteId: () => undefined,
    findDuplicateNoteIds: async (expression) => {
      duplicateLookupExpressions.push(expression);
      return [18, 7, 30];
    },
    trackLastAddedDuplicateNoteIds: (noteId, duplicateNoteIds) => {
      trackedDuplicates.push({ noteId, duplicateNoteIds });
    },
  });

  const created = await service.createSentenceCard('重複文', 0, 1);

  assert.equal(created, true);
  assert.deepEqual(duplicateLookupExpressions, ['重複文']);
  assert.deepEqual(trackedDuplicates, [{ noteId: 42, duplicateNoteIds: [7, 18, 30] }]);
});

import assert from 'node:assert/strict';
import test from 'node:test';

import { CardCreationService } from './card-creation';
import type { AnkiConnectConfig } from '../types';

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

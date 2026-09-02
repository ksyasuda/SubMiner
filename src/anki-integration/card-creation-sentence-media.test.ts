import assert from 'node:assert/strict';
import test from 'node:test';

import { CardCreationService } from './card-creation';
import type { AnkiConnectConfig } from '../types/anki';

type CardCreationDeps = ConstructorParameters<typeof CardCreationService>[0];

test('sentence card writes generated audio only to sentence audio field', async () => {
  const addedFields: Record<string, string>[] = [];
  const updatedFields: Record<string, string>[] = [];
  const storedMedia: string[] = [];
  const requestedProperties: string[] = [];
  const audioVolumeScales: Array<number | undefined> = [];
  const audioRanges: Array<{ start: number; end: number; padding: number | undefined }> = [];

  const deps: CardCreationDeps = {
    getConfig: () =>
      ({
        deck: 'Mining',
        fields: {
          word: 'Expression',
          sentence: 'Sentence',
          audio: 'ExpressionAudio',
          translation: 'SelectionText',
        },
        media: {
          generateAudio: true,
          generateImage: false,
          mirrorMpvVolume: true,
          maxMediaDuration: 30,
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
        currentSubStart: 12,
        currentSubEnd: 14,
        currentTimePos: 13,
        currentAudioStreamIndex: 0,
        requestProperty: async (name: string) => {
          requestedProperties.push(name);
          return 40;
        },
      }) as never,
    client: {
      addNote: async (_deck, _modelName, fields) => {
        addedFields.push(fields);
        return 42;
      },
      addTags: async () => undefined,
      notesInfo: async () => [
        {
          noteId: 42,
          fields: {
            Expression: { value: '字幕' },
            Sentence: { value: '字幕' },
            SelectionText: { value: 'Subtitle' },
            ExpressionAudio: { value: '' },
            SentenceAudio: { value: '' },
          },
        },
      ],
      updateNoteFields: async (_noteId, fields) => {
        updatedFields.push(fields);
      },
      storeMediaFile: async (filename) => {
        storedMedia.push(filename);
      },
      findNotes: async () => [],
      retrieveMediaFile: async () => '',
      deleteNotes: async () => undefined,
    },
    mediaGenerator: {
      generateAudio: async (
        _path,
        startTime,
        endTime,
        audioPadding,
        _audioStreamIndex,
        _normalizeAudio,
        volumeScale,
      ) => {
        audioRanges.push({ start: startTime, end: endTime, padding: audioPadding });
        audioVolumeScales.push(volumeScale);
        return Buffer.from('audio');
      },
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
    extractFields: () => ({}),
    processSentence: (sentence) => sentence,
    setCardTypeFields: () => undefined,
    mergeFieldValue: (_existing, newValue) => newValue,
    formatMiscInfoPattern: () => '',
    getEffectiveSentenceCardConfig: () => ({
      model: 'Sentence',
      sentenceField: 'Sentence',
      audioField: 'SentenceAudio',
      lapisEnabled: true,
      kikuEnabled: false,
      fieldGroupingMode: 'disabled',
    }),
    getFallbackDurationSeconds: () => 10,
    appendKnownWordsFromNoteInfo: () => undefined,
    removeKnownWordNote: () => undefined,
    isUpdateInProgress: () => false,
    setUpdateInProgress: () => undefined,
    trackLastAddedNoteId: () => undefined,
    reviewMediaTiming: async () => ({ action: 'confirm', startTime: 11.4, endTime: 14.2 }),
  };

  const service = new CardCreationService(deps);
  const created = await service.createSentenceCard('字幕', 12, 14, 'Subtitle');

  assert.equal(created, true);
  assert.deepEqual(addedFields[0], {
    Sentence: '字幕',
    SelectionText: 'Subtitle',
    IsSentenceCard: 'x',
    Expression: '字幕',
  });
  assert.equal(storedMedia.length, 1);
  assert.deepEqual(requestedProperties, ['volume']);
  assert.deepEqual(audioVolumeScales, [0.4 ** 3]);
  assert.deepEqual(audioRanges, [{ start: 11.4, end: 14.2, padding: 0 }]);
  const mediaUpdate = updatedFields.find((fields) => 'SentenceAudio' in fields);
  assert.equal(mediaUpdate?.SentenceAudio, `[sound:${storedMedia[0]}]`);
  assert.equal('ExpressionAudio' in mediaUpdate!, false);

  deps.reviewMediaTiming = async () => ({ action: 'discard' });
  assert.equal(await service.createSentenceCard('作らない', 20, 22), false);
  assert.equal(addedFields.length, 1);

  deps.reviewMediaTiming = async () => ({ action: 'skip-media' });
  assert.equal(await service.createSentenceCard('メディアなし', 30, 32), true);
  assert.equal(addedFields.length, 2);
  assert.equal(storedMedia.length, 1);
  assert.deepEqual(audioRanges, [{ start: 11.4, end: 14.2, padding: 0 }]);
  assert.deepEqual(requestedProperties, ['volume']);
});

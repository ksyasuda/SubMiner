import assert from 'node:assert/strict';
import test from 'node:test';

import { CardCreationService } from './card-creation';
import {
  type CardCreationDeps,
  createAnkiConfig,
  createCardCreationDeps,
  createFakeClient,
  createFakeMediaGenerator,
} from './card-creation-test-utils';

test('sentence card writes generated audio only to sentence audio field', async () => {
  const addedFields: Record<string, string>[] = [];
  const updatedFields: Record<string, string>[] = [];
  const storedMedia: string[] = [];
  const requestedProperties: string[] = [];
  const audioVolumeScales: Array<number | undefined> = [];
  const audioRanges: Array<{ start: number; end: number; padding: number | undefined }> = [];
  let review: Awaited<ReturnType<NonNullable<CardCreationDeps['reviewMediaTiming']>>> = {
    action: 'confirm',
    startTime: 11.4,
    endTime: 14.2,
  };

  const service = new CardCreationService(
    createCardCreationDeps({
      getConfig: () =>
        createAnkiConfig({
          fields: {
            word: 'Expression',
            audio: 'ExpressionAudio',
            translation: 'SelectionText',
          },
          media: { generateAudio: true, mirrorMpvVolume: true, maxMediaDuration: 30 },
        }),
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
      client: createFakeClient({
        addNote: async (_deck, _modelName, fields) => {
          addedFields.push(fields);
          return 42;
        },
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
      }),
      mediaGenerator: createFakeMediaGenerator({
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
      }),
      getEffectiveSentenceCardConfig: () => ({
        model: 'Sentence',
        sentenceField: 'Sentence',
        audioField: 'SentenceAudio',
        lapisEnabled: true,
        kikuEnabled: false,
        fieldGroupingMode: 'disabled',
      }),
      reviewMediaTiming: async () => review,
    }),
  );

  assert.equal(await service.createSentenceCard('字幕', 12, 14, 'Subtitle'), true);
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

  review = { action: 'discard' };
  assert.equal(await service.createSentenceCard('作らない', 20, 22), false);
  assert.equal(addedFields.length, 1);

  review = { action: 'skip-media' };
  assert.equal(await service.createSentenceCard('メディアなし', 30, 32), true);
  assert.equal(addedFields.length, 2);
  assert.equal(storedMedia.length, 1);
  assert.deepEqual(audioRanges, [{ start: 11.4, end: 14.2, padding: 0 }]);
  assert.deepEqual(requestedProperties, ['volume']);
});

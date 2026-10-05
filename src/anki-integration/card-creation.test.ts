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
import { toMpvEdlValue } from './mpv-edl-test-utils';
import type { MediaInput } from '../media-generator';

const YOUTUBE_URL = 'https://www.youtube.com/watch?v=abc123';

const mediaInputPath = (mediaInput: MediaInput): string =>
  typeof mediaInput === 'string' ? mediaInput : mediaInput.path;

test('CardCreationService counts locally created sentence cards', async () => {
  const minedCards: Array<{ count: number; noteIds?: number[] }> = [];
  const service = new CardCreationService(
    createCardCreationDeps({
      recordCardsMinedCallback: (count, noteIds) => {
        minedCards.push({ count, noteIds });
      },
    }),
  );

  assert.equal(await service.createSentenceCard('テスト', 0, 1), true);
  assert.deepEqual(minedCards, [{ count: 1, noteIds: [42] }]);
});

const THROWING_HOOKS: Array<{
  name: string;
  hook: Partial<CardCreationDeps>;
}> = [
  {
    name: 'trackLastAddedNoteId',
    hook: {
      trackLastAddedNoteId: () => {
        throw new Error('track failed');
      },
    },
  },
  {
    name: 'recordCardsMinedCallback',
    hook: {
      recordCardsMinedCallback: () => {
        throw new Error('record failed');
      },
    },
  },
];

for (const c of THROWING_HOOKS) {
  test(`CardCreationService keeps updating after ${c.name} throws`, async () => {
    const calls = { notesInfo: 0, updateNoteFields: 0 };
    const service = new CardCreationService(
      createCardCreationDeps({
        client: createFakeClient({
          notesInfo: async () => {
            calls.notesInfo += 1;
            return [{ noteId: 42, fields: { Sentence: { value: 'existing' } } }];
          },
          updateNoteFields: async () => {
            calls.updateNoteFields += 1;
          },
        }),
        setCardTypeFields: (updatedFields) => {
          updatedFields.CardType = 'sentence';
        },
        ...c.hook,
      }),
    );

    assert.equal(await service.createSentenceCard('テスト', 0, 1), true);
    assert.equal(calls.notesInfo, 1);
    assert.equal(calls.updateNoteFields, 1);
  });
}

test('CardCreationService uses stream-open-filename for remote media generation', async () => {
  let reviewing = false;
  const audioRanges: number[][] = [];
  const imageTimes: number[] = [];
  const audioPaths: string[] = [];
  const imagePaths: string[] = [];
  const audioUrl = 'https://audio.example/videoplayback?mime=audio%2Fwebm';
  const videoUrl = 'https://video.example/videoplayback?mime=video%2Fmp4';
  const edlSource = [
    `edl://!new_stream;!no_clip;!no_chapters;${toMpvEdlValue(audioUrl)}`,
    `!new_stream;!no_clip;!no_chapters;${toMpvEdlValue(videoUrl)}`,
    '!global_tags,title=test',
  ].join(';');

  const service = new CardCreationService(
    createCardCreationDeps({
      getConfig: () =>
        createAnkiConfig({
          fields: { image: 'Picture' },
          media: { generateAudio: true, generateImage: true, imageFormat: 'jpg' },
        }),
      reviewMediaTiming: async () =>
        reviewing
          ? { action: 'confirm', startTime: 0.2, endTime: 0.8, screenshotTime: 3.125 }
          : { action: 'use-original' },
      getMpvClient: () =>
        ({
          currentVideoPath: YOUTUBE_URL,
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
      client: createFakeClient({
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
        findNotes: async () => [42],
      }),
      mediaGenerator: createFakeMediaGenerator({
        generateAudio: async (path, start, end, padding) => {
          audioRanges.push([start, end, padding ?? -1]);
          audioPaths.push(mediaInputPath(path));
          return Buffer.from('audio');
        },
        generateScreenshot: async (path, timestamp) => {
          imageTimes.push(timestamp);
          imagePaths.push(mediaInputPath(path));
          return Buffer.from('image');
        },
      }),
    }),
  );

  assert.equal(await service.createSentenceCard('テスト', 0, 1), true);
  assert.deepEqual(audioPaths, [audioUrl]);
  assert.deepEqual(imagePaths, [videoUrl]);

  reviewing = true;
  assert.equal(await service.createSentenceCard('テスト', 0, 1), true);
  assert.deepEqual(audioRanges.at(-1), [0.2, 0.8, 0]);
  assert.equal(imageTimes.at(-1), 3.125);

  await service.markLastCardAsAudioCard();
  assert.equal(imageTimes.length, 3);
  assert.deepEqual(audioRanges.at(-1), [0.2, 0.8, 0]);
  assert.equal(imageTimes.at(-1), 3.125);
});

test('CardCreationService does not use mpv stream indexes for ready cached YouTube media', async () => {
  const audioCalls: Array<{ path: string; audioStreamIndex?: number }> = [];

  const service = new CardCreationService(
    createCardCreationDeps({
      getConfig: () =>
        createAnkiConfig({
          fields: { image: 'Picture' },
          media: { generateAudio: true, imageFormat: 'jpg' },
        }),
      getMpvClient: () =>
        ({
          currentVideoPath: YOUTUBE_URL,
          currentSubText: '字幕',
          currentSubStart: 10,
          currentSubEnd: 12,
          currentTimePos: 11,
          currentAudioStreamIndex: 0,
        }) as never,
      getCachedMediaPath: async () => '/tmp/subminer-youtube-media-cache/media.mkv',
      shouldRequireRemoteMediaCache: () => true,
      client: createFakeClient({
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
      }),
      mediaGenerator: createFakeMediaGenerator({
        generateAudio: async (path, _startTime, _endTime, _padding, audioStreamIndex) => {
          audioCalls.push({ path: mediaInputPath(path), audioStreamIndex });
          return Buffer.from('audio');
        },
      }),
    }),
  );

  assert.equal(await service.createSentenceCard('テスト', 10, 12), true);
  assert.deepEqual(audioCalls, [
    { path: '/tmp/subminer-youtube-media-cache/media.mkv', audioStreamIndex: undefined },
  ]);
});

test('CardCreationService queues YouTube media when required cache is not ready', async () => {
  const mediaCalls: string[] = [];
  const updates: Array<{ noteId: number; fields: Record<string, string> }> = [];
  const queuedUpdates: Array<{
    sourceUrl: string;
    noteId: number;
    startTime: number;
    endTime: number;
    label: string | number;
    audioFieldName?: string;
    imageFieldName?: string;
    miscInfoFieldName?: string;
    generateAudio: boolean;
    generateImage: boolean;
    volumeScale?: number;
  }> = [];
  let streamRequests = 0;

  const service = new CardCreationService(
    createCardCreationDeps({
      getConfig: () =>
        createAnkiConfig({
          fields: { image: 'Picture', miscInfo: 'MiscInfo' },
          media: { generateAudio: true, generateImage: true, imageFormat: 'jpg' },
        }),
      getMpvClient: () =>
        ({
          currentVideoPath: YOUTUBE_URL,
          currentSubText: '字幕',
          currentSubStart: 10,
          currentSubEnd: 12,
          currentTimePos: 11,
          currentAudioStreamIndex: 2,
          requestProperty: async (name: string) => {
            if (name === 'volume') return 35;
            streamRequests += 1;
            return 'https://rr1---sn.example.googlevideo.com/videoplayback?id=123';
          },
        }) as never,
      getCachedMediaPath: async () => null,
      shouldRequireRemoteMediaCache: () => true,
      queuePendingYoutubeMediaUpdate: (job) => {
        queuedUpdates.push(job);
      },
      client: createFakeClient({
        notesInfo: async () => [
          {
            noteId: 42,
            fields: {
              Sentence: { value: '' },
              SentenceAudio: { value: '' },
              Picture: { value: '' },
              MiscInfo: { value: '' },
            },
          },
        ],
        updateNoteFields: async (noteId, fields) => {
          updates.push({ noteId, fields });
        },
      }),
      mediaGenerator: createFakeMediaGenerator({
        generateAudio: async () => {
          mediaCalls.push('audio');
          return Buffer.from('audio');
        },
        generateScreenshot: async () => {
          mediaCalls.push('image');
          return Buffer.from('image');
        },
      }),
    }),
  );

  assert.equal(await service.createSentenceCard('テスト', 10, 12), true);
  assert.equal(streamRequests, 0);
  assert.deepEqual(mediaCalls, []);
  assert.deepEqual(queuedUpdates, [
    {
      sourceUrl: YOUTUBE_URL,
      noteId: 42,
      startTime: 10,
      endTime: 12,
      label: 'テスト',
      audioFieldName: 'SentenceAudio',
      imageFieldName: 'Picture',
      miscInfoFieldName: 'MiscInfo',
      generateAudio: true,
      generateImage: true,
      volumeScale: 0.35 ** 3,
    },
  ]);
  assert.deepEqual(updates, []);
});

const DUPLICATE_TRACKING_CASES = [
  {
    name: 'tracks sorted, deduplicated pre-add duplicate note ids for kiku sentence cards',
    sentence: '重複文',
    lookupResult: [18, 7, 30, 7],
    expected: [{ noteId: 42, duplicateNoteIds: [7, 18, 30] }],
  },
  {
    name: 'does not track duplicate ids when pre-add lookup returns none',
    sentence: '重複なし',
    lookupResult: [],
    expected: [],
  },
];

for (const c of DUPLICATE_TRACKING_CASES) {
  test(`CardCreationService ${c.name}`, async () => {
    const trackedDuplicates: Array<{ noteId: number; duplicateNoteIds: number[] }> = [];
    const lookedUp: string[] = [];
    const service = new CardCreationService(
      createCardCreationDeps({
        getConfig: () => createAnkiConfig({ fields: { word: 'Expression' } }),
        getEffectiveSentenceCardConfig: () => ({
          model: 'Sentence',
          sentenceField: 'Sentence',
          audioField: 'SentenceAudio',
          lapisEnabled: false,
          kikuEnabled: true,
          fieldGroupingMode: 'manual',
        }),
        findDuplicateNoteIds: async (expression) => {
          lookedUp.push(expression);
          return c.lookupResult;
        },
        trackLastAddedDuplicateNoteIds: (noteId, duplicateNoteIds) => {
          trackedDuplicates.push({ noteId, duplicateNoteIds });
        },
      }),
    );

    assert.equal(await service.createSentenceCard(c.sentence, 0, 1), true);
    assert.deepEqual(lookedUp, [c.sentence]);
    assert.deepEqual(trackedDuplicates, c.expected);
  });
}

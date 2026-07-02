import assert from 'node:assert/strict';
import test from 'node:test';

import { CardCreationService } from './card-creation';
import type { MediaInput } from '../media-generator';
import type { AnkiConnectConfig } from '../types/anki';

function toMpvEdlValue(value: string): string {
  return `%${Buffer.byteLength(value, 'utf8')}%${value}`;
}

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
  const recordMediaPath = (mediaInput: MediaInput): string =>
    typeof mediaInput === 'string' ? mediaInput : mediaInput.path;
  const audioUrl = 'https://audio.example/videoplayback?mime=audio%2Fwebm';
  const videoUrl = 'https://video.example/videoplayback?mime=video%2Fmp4';
  const edlSource = [
    `edl://!new_stream;!no_clip;!no_chapters;${toMpvEdlValue(audioUrl)}`,
    `!new_stream;!no_clip;!no_chapters;${toMpvEdlValue(videoUrl)}`,
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
        audioPaths.push(recordMediaPath(path));
        return Buffer.from('audio');
      },
      generateScreenshot: async (path) => {
        imagePaths.push(recordMediaPath(path));
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
  assert.deepEqual(audioPaths, [audioUrl]);
  assert.deepEqual(imagePaths, [videoUrl]);
});

test('CardCreationService does not use mpv stream indexes for ready cached YouTube media', async () => {
  const audioCalls: Array<{ path: string; audioStreamIndex?: number }> = [];

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
          generateImage: false,
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
        currentSubStart: 10,
        currentSubEnd: 12,
        currentTimePos: 11,
        currentAudioStreamIndex: 0,
      }) as never,
    getCachedMediaPath: async () => '/tmp/subminer-youtube-media-cache/media.mkv',
    shouldRequireRemoteMediaCache: () => true,
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
      generateAudio: async (path, _startTime, _endTime, _padding, audioStreamIndex) => {
        audioCalls.push({ path: typeof path === 'string' ? path : path.path, audioStreamIndex });
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

  const created = await service.createSentenceCard('テスト', 10, 12);

  assert.equal(created, true);
  assert.deepEqual(audioCalls, [
    {
      path: '/tmp/subminer-youtube-media-cache/media.mkv',
      audioStreamIndex: undefined,
    },
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
  }> = [];
  let streamRequests = 0;

  const service = new CardCreationService({
    getConfig: () =>
      ({
        deck: 'Mining',
        fields: {
          sentence: 'Sentence',
          audio: 'SentenceAudio',
          image: 'Picture',
          miscInfo: 'MiscInfo',
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
        currentSubStart: 10,
        currentSubEnd: 12,
        currentTimePos: 11,
        currentAudioStreamIndex: 2,
        requestProperty: async () => {
          streamRequests += 1;
          return 'https://rr1---sn.example.googlevideo.com/videoplayback?id=123';
        },
      }) as never,
    getCachedMediaPath: async () => null,
    shouldRequireRemoteMediaCache: () => true,
    queuePendingYoutubeMediaUpdate: (job) => {
      queuedUpdates.push(job);
    },
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
            MiscInfo: { value: '' },
          },
        },
      ],
      updateNoteFields: async (noteId, fields) => {
        updates.push({ noteId, fields });
      },
      storeMediaFile: async () => undefined,
      findNotes: async () => [],
      retrieveMediaFile: async () => '',
    },
    mediaGenerator: {
      generateAudio: async () => {
        mediaCalls.push('audio');
        return Buffer.from('audio');
      },
      generateScreenshot: async () => {
        mediaCalls.push('image');
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

  const created = await service.createSentenceCard('テスト', 10, 12);

  assert.equal(created, true);
  assert.equal(streamRequests, 0);
  assert.deepEqual(mediaCalls, []);
  assert.deepEqual(queuedUpdates, [
    {
      sourceUrl: 'https://www.youtube.com/watch?v=abc123',
      noteId: 42,
      startTime: 10,
      endTime: 12,
      label: 'テスト',
      audioFieldName: 'SentenceAudio',
      imageFieldName: 'Picture',
      miscInfoFieldName: 'MiscInfo',
      generateAudio: true,
      generateImage: true,
    },
  ]);
  assert.deepEqual(updates, []);
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
      return [18, 7, 30, 7];
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

test('CardCreationService does not track duplicate ids when pre-add lookup returns none', async () => {
  const trackedDuplicates: Array<{ noteId: number; duplicateNoteIds: number[] }> = [];

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
    findDuplicateNoteIds: async () => [],
    trackLastAddedDuplicateNoteIds: (noteId, duplicateNoteIds) => {
      trackedDuplicates.push({ noteId, duplicateNoteIds });
    },
  });

  const created = await service.createSentenceCard('重複なし', 0, 1);

  assert.equal(created, true);
  assert.deepEqual(trackedDuplicates, []);
});

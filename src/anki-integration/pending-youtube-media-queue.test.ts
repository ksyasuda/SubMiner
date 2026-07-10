import assert from 'node:assert/strict';
import test from 'node:test';

import type { AnkiConnectConfig } from '../types/anki';
import {
  PendingYoutubeMediaQueue,
  type PendingYoutubeMediaQueueDeps,
} from './pending-youtube-media-queue';

function createDeps(
  overrides: Partial<PendingYoutubeMediaQueueDeps> = {},
): PendingYoutubeMediaQueueDeps {
  const warnings: unknown[][] = [];
  const deps: PendingYoutubeMediaQueueDeps & { warnings: unknown[][] } = {
    client: {
      notesInfo: async () => [],
      updateNoteFields: async () => {},
      storeMediaFile: async () => {},
    },
    mediaGenerator: {
      generateAudio: async () => Buffer.from('audio'),
      generateScreenshot: async () => Buffer.from('image'),
      generateAnimatedImage: async () => Buffer.from('image'),
    },
    getConfig: () =>
      ({
        media: { generateAudio: true, generateImage: true },
        fields: {},
      }) as AnkiConnectConfig,
    getCurrentVideoPath: () => 'https://www.youtube.com/watch?v=abc123',
    getCachedMediaPath: async () => null,
    shouldRequireRemoteMediaCache: () => true,
    getSubtitleMediaRange: () => ({ startTime: 1, endTime: 2 }),
    getResolvedSentenceAudioFieldName: () => 'SentenceAudio',
    resolveConfiguredFieldName: () => 'Picture',
    mergeFieldValue: (_existing, newValue) => newValue,
    getAnimatedImageLeadInSeconds: async () => 0,
    getMpvVolumeScale: async () => 1,
    generateAudioFilename: () => 'audio.mp3',
    generateImageFilename: () => 'image.webp',
    formatMiscInfoPatternForMediaPath: () => '',
    showStatusNotification: () => {},
    showNotification: async () => {},
    logInfo: () => {},
    logWarn: (...args) => {
      warnings.push(args);
    },
    logError: () => {},
    warnings,
    ...overrides,
  };
  return deps;
}

test('PendingYoutubeMediaQueue treats cache lookup failures as an immediate generation fallback', async () => {
  const deps = createDeps({
    getCachedMediaPath: async () => {
      throw new Error('cache unavailable');
    },
  });
  const queue = new PendingYoutubeMediaQueue(deps);

  const queued = await queue.queueFromNote({
    noteId: 42,
    noteInfo: { noteId: 42, fields: {} },
    label: 'demo',
  });

  assert.equal(queued, false);
  assert.deepEqual((deps as typeof deps & { warnings: unknown[][] }).warnings, [
    [
      'Failed to read YouTube cache state; falling back to immediate media generation:',
      'cache unavailable',
    ],
  ]);
});

test('PendingYoutubeMediaQueue drains matching queued jobs when the cache download fails', async () => {
  const statusMessages: string[] = [];
  const notifications: Array<{ noteId: number; label: string | number; suffix?: string }> = [];
  const updatedNotes: number[] = [];
  const deps = createDeps({
    client: {
      notesInfo: async (noteIds) =>
        noteIds.map((noteId) => ({
          noteId,
          fields: {
            SentenceAudio: { value: '' },
            Picture: { value: '' },
          },
        })),
      updateNoteFields: async (noteId) => {
        updatedNotes.push(noteId);
      },
      storeMediaFile: async () => {},
    },
    showStatusNotification: (message) => {
      statusMessages.push(message);
    },
    showNotification: async (noteId, label, suffix) => {
      notifications.push({ noteId, label, suffix });
    },
  });
  const queue = new PendingYoutubeMediaQueue(deps);

  queue.enqueue({
    sourceUrl: 'https://www.youtube.com/watch?v=abc123',
    noteId: 42,
    startTime: 1,
    endTime: 2,
    label: 'queued',
    generateAudio: true,
    generateImage: true,
  });

  await queue.handleFailed('https://youtu.be/abc123');
  await queue.handleReady('https://youtu.be/abc123', '/tmp/media.mkv');

  assert.deepEqual(updatedNotes, []);
  assert.deepEqual(notifications, [{ noteId: 42, label: 'queued', suffix: 'media cache failed' }]);
  assert.equal(
    statusMessages.includes('YouTube media cache failed. Media was not added to 1 queued card.'),
    true,
  );
});

test('PendingYoutubeMediaQueue defaults missing media flags to enabled when queuing from notes', async () => {
  const updatedNotes: Array<{ noteId: number; fields: Record<string, string> }> = [];
  const storedMedia: string[] = [];
  const deps = createDeps({
    client: {
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
    },
    getConfig: () => ({ media: {}, fields: { image: 'Picture' } }) as AnkiConnectConfig,
  });
  const queue = new PendingYoutubeMediaQueue(deps);

  const queued = await queue.queueFromNote({
    noteId: 42,
    noteInfo: { noteId: 42, fields: {} },
    label: 'demo',
  });
  await queue.handleReady('https://youtu.be/abc123', '/tmp/media.mkv');

  assert.equal(queued, true);
  assert.equal(updatedNotes.length, 1);
  assert.equal(storedMedia.length, 2);
  assert.match(updatedNotes[0]?.fields.SentenceAudio ?? '', /^\[sound:audio\.mp3\]$/);
  assert.match(updatedNotes[0]?.fields.Picture ?? '', /^<img src="image\.webp">$/);
});

test('PendingYoutubeMediaQueue only announces a download once per source while jobs collect', () => {
  const statusMessages: string[] = [];
  const deps = createDeps({
    showStatusNotification: (message) => {
      statusMessages.push(message);
    },
  });
  const queue = new PendingYoutubeMediaQueue(deps);

  queue.enqueue({
    sourceUrl: 'https://youtu.be/abc123',
    noteId: 1,
    startTime: 1,
    endTime: 2,
    label: 'first',
    generateAudio: true,
    generateImage: false,
  });
  queue.enqueue({
    sourceUrl: 'https://www.youtube.com/watch?v=abc123',
    noteId: 2,
    startTime: 3,
    endTime: 4,
    label: 'second',
    generateAudio: false,
    generateImage: true,
  });

  assert.equal(statusMessages.length, 1);
});

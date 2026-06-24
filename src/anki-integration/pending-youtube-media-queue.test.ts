import assert from 'node:assert/strict';
import test from 'node:test';

import type { AnkiConnectConfig } from '../types/anki';
import { PendingYoutubeMediaQueue, type PendingYoutubeMediaQueueDeps } from './pending-youtube-media-queue';

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

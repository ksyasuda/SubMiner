import assert from 'node:assert/strict';
import test from 'node:test';
import type { SubtitlePrefetchService } from '../../core/services/subtitle-prefetch';
import type { SubtitleCue } from '../../types';
import { createSubtitlePrefetchInitController } from './subtitle-prefetch-init';

function createDeferred<T>(): {
  promise: Promise<T>;
  resolve: (value: T) => void;
  reject: (error: unknown) => void;
} {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

function flushMicrotasks(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

test('latest subtitle prefetch init wins over stale async loads', async () => {
  const loads = new Map<string, ReturnType<typeof createDeferred<string>>>();
  const started: string[] = [];
  const stopped: string[] = [];
  let currentService: SubtitlePrefetchService | null = null;

  const controller = createSubtitlePrefetchInitController({
    getCurrentService: () => currentService,
    setCurrentService: (service) => {
      currentService = service;
    },
    loadSubtitleSourceText: async (source) => {
      const deferred = createDeferred<string>();
      loads.set(source, deferred);
      return await deferred.promise;
    },
    parseSubtitleCues: (_content, filename): SubtitleCue[] => [
      { startTime: 0, endTime: 1, text: filename },
    ],
    createSubtitlePrefetchService: ({ cues }) => ({
      start: () => {
        started.push(cues[0]!.text);
      },
      stop: () => {
        stopped.push(cues[0]!.text);
      },
      onSeek: () => {},
      pause: () => {},
      resume: () => {},
    }),
    tokenizeSubtitle: async () => null,
    preCacheTokenization: () => {},
    logInfo: () => {},
    logWarn: () => {},
  });

  const firstInit = controller.initSubtitlePrefetch('old.ass', 1);
  const secondInit = controller.initSubtitlePrefetch('new.ass', 2);

  loads.get('new.ass')!.resolve('new');
  await flushMicrotasks();

  assert.deepEqual(started, ['new.ass']);

  loads.get('old.ass')!.resolve('old');
  await Promise.all([firstInit, secondInit]);

  assert.deepEqual(started, ['new.ass']);
  assert.deepEqual(stopped, []);
});

test('cancelPendingInit prevents an in-flight load from attaching a stale service', async () => {
  const deferred = createDeferred<string>();
  let currentService: SubtitlePrefetchService | null = null;
  const started: string[] = [];

  const controller = createSubtitlePrefetchInitController({
    getCurrentService: () => currentService,
    setCurrentService: (service) => {
      currentService = service;
    },
    loadSubtitleSourceText: async () => await deferred.promise,
    parseSubtitleCues: (_content, filename): SubtitleCue[] => [
      { startTime: 0, endTime: 1, text: filename },
    ],
    createSubtitlePrefetchService: ({ cues }) => ({
      start: () => {
        started.push(cues[0]!.text);
      },
      stop: () => {},
      onSeek: () => {},
      pause: () => {},
      resume: () => {},
    }),
    tokenizeSubtitle: async () => null,
    preCacheTokenization: () => {},
    logInfo: () => {},
    logWarn: () => {},
  });

  const initPromise = controller.initSubtitlePrefetch('stale.ass', 1);
  controller.cancelPendingInit();
  deferred.resolve('stale');
  await initPromise;

  assert.equal(currentService, null);
  assert.deepEqual(started, []);
});

test('subtitle prefetch init publishes parsed cues and clears them on cancel', async () => {
  const deferred = createDeferred<string>();
  let currentService: SubtitlePrefetchService | null = null;
  const cueUpdates: Array<SubtitleCue[] | null> = [];

  const controller = createSubtitlePrefetchInitController({
    getCurrentService: () => currentService,
    setCurrentService: (service) => {
      currentService = service;
    },
    loadSubtitleSourceText: async () => await deferred.promise,
    parseSubtitleCues: () => [
      { startTime: 1, endTime: 2, text: 'first' },
      { startTime: 3, endTime: 4, text: 'second' },
    ],
    createSubtitlePrefetchService: () => ({
      start: () => {},
      stop: () => {},
      onSeek: () => {},
      pause: () => {},
      resume: () => {},
    }),
    tokenizeSubtitle: async () => null,
    preCacheTokenization: () => {},
    logInfo: () => {},
    logWarn: () => {},
    onParsedSubtitleCuesChanged: (cues) => {
      cueUpdates.push(cues);
    },
  });

  const initPromise = controller.initSubtitlePrefetch('episode.ass', 12);
  deferred.resolve('content');
  await initPromise;

  controller.cancelPendingInit();

  assert.deepEqual(cueUpdates, [
    [
      { startTime: 1, endTime: 2, text: 'first' },
      { startTime: 3, endTime: 4, text: 'second' },
    ],
    null,
  ]);
});

test('subtitle prefetch init publishes the provided stable source key instead of the load path', async () => {
  const deferred = createDeferred<string>();
  let currentService: SubtitlePrefetchService | null = null;
  const sourceUpdates: Array<string | null> = [];

  const controller = createSubtitlePrefetchInitController({
    getCurrentService: () => currentService,
    setCurrentService: (service) => {
      currentService = service;
    },
    loadSubtitleSourceText: async () => await deferred.promise,
    parseSubtitleCues: () => [{ startTime: 1, endTime: 2, text: 'first' }],
    createSubtitlePrefetchService: () => ({
      start: () => {},
      stop: () => {},
      onSeek: () => {},
      pause: () => {},
      resume: () => {},
    }),
    tokenizeSubtitle: async () => null,
    preCacheTokenization: () => {},
    logInfo: () => {},
    logWarn: () => {},
    onParsedSubtitleCuesChanged: (_cues, source) => {
      sourceUpdates.push(source);
    },
  });

  const initPromise = controller.initSubtitlePrefetch(
    '/tmp/subminer-sidebar-123/track_7.ass',
    12,
    'internal:/media/episode01.mkv:track:3:ff:7',
  );
  deferred.resolve('content');
  await initPromise;

  assert.deepEqual(sourceUpdates, ['internal:/media/episode01.mkv:track:3:ff:7']);
});

test('subtitle prefetch init clears parsed cues when initialization fails', async () => {
  const cueUpdates: Array<SubtitleCue[] | null> = [];
  let currentService: SubtitlePrefetchService | null = null;

  const controller = createSubtitlePrefetchInitController({
    getCurrentService: () => currentService,
    setCurrentService: (service) => {
      currentService = service;
    },
    loadSubtitleSourceText: async () => {
      throw new Error('boom');
    },
    parseSubtitleCues: () => [{ startTime: 1, endTime: 2, text: 'first' }],
    createSubtitlePrefetchService: () => ({
      start: () => {},
      stop: () => {},
      onSeek: () => {},
      pause: () => {},
      resume: () => {},
    }),
    tokenizeSubtitle: async () => null,
    preCacheTokenization: () => {},
    logInfo: () => {},
    logWarn: () => {},
    onParsedSubtitleCuesChanged: (cues) => {
      cueUpdates.push(cues);
    },
  });

  await controller.initSubtitlePrefetch('episode.ass', 12);

  assert.deepEqual(cueUpdates, [null]);
});

test('subtitle prefetch init logs a warning when the source parses to zero cues', async () => {
  const warnings: string[] = [];
  const controller = createSubtitlePrefetchInitController({
    getCurrentService: () => null,
    setCurrentService: () => {},
    loadSubtitleSourceText: async () => 'not really subtitles',
    parseSubtitleCues: (): SubtitleCue[] => [],
    createSubtitlePrefetchService: () => {
      throw new Error('should not create a service without cues');
    },
    tokenizeSubtitle: async () => null,
    preCacheTokenization: () => {},
    logInfo: () => {},
    logWarn: (message) => warnings.push(message),
  });

  await controller.initSubtitlePrefetch('/tmp/broken.ass', 0);

  assert.equal(warnings.length, 1);
  assert.match(warnings[0]!, /\[subtitle-prefetch\].*0 cues.*\/tmp\/broken\.ass/);
});

function createTrackingController(loadContent: () => string) {
  let currentService: SubtitlePrefetchService | null = null;
  const events: string[] = [];
  const cueUpdates: Array<string | null> = [];
  let serviceCount = 0;

  const controller = createSubtitlePrefetchInitController({
    getCurrentService: () => currentService,
    setCurrentService: (service) => {
      currentService = service;
    },
    loadSubtitleSourceText: async () => loadContent(),
    parseSubtitleCues: (content): SubtitleCue[] => [{ startTime: 0, endTime: 1, text: content }],
    createSubtitlePrefetchService: () => {
      const id = ++serviceCount;
      return {
        start: () => events.push(`start:${id}`),
        stop: () => events.push(`stop:${id}`),
        onSeek: () => {},
        pause: () => {},
        resume: () => {},
      };
    },
    tokenizeSubtitle: async () => null,
    preCacheTokenization: () => {},
    logInfo: () => {},
    logWarn: () => {},
    onParsedSubtitleCuesChanged: (cues, source) => {
      cueUpdates.push(cues ? `${source}:${cues[0]!.text}` : null);
    },
  });

  return { controller, events, cueUpdates };
}

test('re-initializing an unchanged source keeps the running service and republishes cues', async () => {
  const { controller, events, cueUpdates } = createTrackingController(() => 'content');

  await controller.initSubtitlePrefetch('track-1.srt', 0);
  await controller.initSubtitlePrefetch('track-1.srt', 3);

  assert.deepEqual(events, ['start:1']);
  assert.deepEqual(cueUpdates, ['track-1.srt:content', 'track-1.srt:content']);
});

test('re-initializing a source whose content changed restarts the service', async () => {
  let content = 'first';
  const { controller, events } = createTrackingController(() => content);

  await controller.initSubtitlePrefetch('track-1.srt', 0);
  content = 'second';
  await controller.initSubtitlePrefetch('track-1.srt', 0);

  assert.deepEqual(events, ['start:1', 'stop:1', 'start:2']);
});

test('subtitle prefetch init warms overlap lines without publishing them as cues', async () => {
  let prefetchedTexts: string[] = [];
  let publishedTexts: string[] = [];
  let currentService: SubtitlePrefetchService | null = null;

  const controller = createSubtitlePrefetchInitController({
    getCurrentService: () => currentService,
    setCurrentService: (service) => {
      currentService = service;
    },
    loadSubtitleSourceText: async () => 'content',
    parseSubtitleCues: (): SubtitleCue[] => [
      { startTime: 0, endTime: 4, text: 'first' },
      { startTime: 2, endTime: 6, text: 'second' },
    ],
    createSubtitlePrefetchService: ({ cues }) => {
      prefetchedTexts = cues.map((cue) => cue.text);
      return {
        start: () => {},
        stop: () => {},
        onSeek: () => {},
        pause: () => {},
        resume: () => {},
      };
    },
    tokenizeSubtitle: async () => null,
    preCacheTokenization: () => {},
    logInfo: () => {},
    logWarn: () => {},
    onParsedSubtitleCuesChanged: (cues) => {
      publishedTexts = cues?.map((cue) => cue.text) ?? [];
    },
  });

  await controller.initSubtitlePrefetch('track.srt', 0);

  assert.deepEqual(prefetchedTexts, ['first', 'second', 'first\n\nsecond']);
  assert.deepEqual(publishedTexts, ['first', 'second']);
});

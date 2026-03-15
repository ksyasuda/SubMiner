import assert from 'node:assert/strict';
import test from 'node:test';
import type { SubtitleCue } from '../../core/services/subtitle-cue-parser';
import type { SubtitlePrefetchService } from '../../core/services/subtitle-prefetch';
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
    isCacheFull: () => false,
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
    isCacheFull: () => false,
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

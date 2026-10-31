import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createMecabTokenizerAndCheckHandler,
  createPrewarmSubtitleDictionariesHandler,
} from './subtitle-tokenization-warmup';

function createDeferred(): {
  promise: Promise<void>;
  resolve: () => void;
} {
  let resolve!: () => void;
  const promise = new Promise<void>((nextResolve) => {
    resolve = nextResolve;
  });
  return { promise, resolve };
}

test('mecab tokenizer check creates tokenizer once and runs availability check', async () => {
  const calls: string[] = [];
  type Tokenizer = { id: string };
  let tokenizer: Tokenizer | null = null;
  const run = createMecabTokenizerAndCheckHandler<Tokenizer>({
    getMecabTokenizer: () => tokenizer,
    setMecabTokenizer: (next) => {
      tokenizer = next;
      calls.push('set');
    },
    createMecabTokenizer: () => {
      calls.push('create');
      return { id: 'mecab' };
    },
    checkAvailability: async () => {
      calls.push('check');
    },
  });

  await run();
  await run();
  assert.deepEqual(calls, ['create', 'set', 'check', 'check']);
});

test('dictionary prewarm runs both dictionary loaders', async () => {
  const calls: string[] = [];
  const prewarm = createPrewarmSubtitleDictionariesHandler({
    ensureJlptDictionaryLookup: async () => {
      calls.push('jlpt');
    },
    ensureFrequencyDictionaryLookup: async () => {
      calls.push('freq');
    },
  });

  await prewarm();
  assert.deepEqual(calls.sort(), ['freq', 'jlpt']);
});

test('dictionary prewarm shows OSD spinner while loading and completion when loaded', async () => {
  const osdMessages: string[] = [];
  const clearedTimers: unknown[] = [];
  let tick: (() => void) | null = null;
  const jlptDeferred = createDeferred();
  const freqDeferred = createDeferred();

  const prewarm = createPrewarmSubtitleDictionariesHandler({
    ensureJlptDictionaryLookup: async () => jlptDeferred.promise,
    ensureFrequencyDictionaryLookup: async () => freqDeferred.promise,
    shouldShowOsdNotification: () => true,
    showMpvOsd: (message) => {
      osdMessages.push(message);
    },
    setInterval: (callback) => {
      tick = callback;
      return 1;
    },
    clearInterval: (timer) => {
      clearedTimers.push(timer);
    },
  });

  const prewarmPromise = prewarm({ showLoadingOsd: true });
  assert.deepEqual(osdMessages, ['Loading subtitle annotations |']);

  if (!tick) {
    throw new Error('expected loading spinner tick callback');
  }
  const tickCallback: () => void = tick;
  tickCallback();
  assert.deepEqual(osdMessages, [
    'Loading subtitle annotations |',
    'Loading subtitle annotations /',
  ]);

  jlptDeferred.resolve();
  freqDeferred.resolve();
  await prewarmPromise;

  assert.deepEqual(osdMessages, [
    'Loading subtitle annotations |',
    'Loading subtitle annotations /',
    'Subtitle annotations loaded',
  ]);
  assert.deepEqual(clearedTimers, [1]);
});

test('dictionary prewarm can show OSD while awaiting background-started load', async () => {
  const osdMessages: string[] = [];
  const jlptDeferred = createDeferred();
  const freqDeferred = createDeferred();

  const prewarm = createPrewarmSubtitleDictionariesHandler({
    ensureJlptDictionaryLookup: async () => jlptDeferred.promise,
    ensureFrequencyDictionaryLookup: async () => freqDeferred.promise,
    shouldShowOsdNotification: () => true,
    showMpvOsd: (message) => {
      osdMessages.push(message);
    },
    setInterval: () => 1,
    clearInterval: () => undefined,
  });

  const backgroundWarmupPromise = prewarm();
  const tokenizationWarmupPromise = prewarm({ showLoadingOsd: true });
  assert.deepEqual(osdMessages, ['Loading subtitle annotations |']);

  jlptDeferred.resolve();
  freqDeferred.resolve();
  await backgroundWarmupPromise;
  await tokenizationWarmupPromise;

  assert.deepEqual(osdMessages, ['Loading subtitle annotations |', 'Subtitle annotations loaded']);
});

test('dictionary prewarm shows OSD when loading indicator is requested even if notification predicate is disabled', async () => {
  const osdMessages: string[] = [];

  const prewarm = createPrewarmSubtitleDictionariesHandler({
    ensureJlptDictionaryLookup: async () => undefined,
    ensureFrequencyDictionaryLookup: async () => undefined,
    shouldShowOsdNotification: () => false,
    showMpvOsd: (message) => {
      osdMessages.push(message);
    },
  });

  await prewarm({ showLoadingOsd: true });

  assert.deepEqual(osdMessages, ['Loading subtitle annotations |', 'Subtitle annotations loaded']);
});

test('dictionary prewarm clears loading OSD timer even if notifications are disabled before completion', async () => {
  const clearedTimers: unknown[] = [];
  const jlptDeferred = createDeferred();
  const freqDeferred = createDeferred();
  let shouldShowNotification = true;

  const prewarm = createPrewarmSubtitleDictionariesHandler({
    ensureJlptDictionaryLookup: async () => jlptDeferred.promise,
    ensureFrequencyDictionaryLookup: async () => freqDeferred.promise,
    shouldShowOsdNotification: () => shouldShowNotification,
    showMpvOsd: () => undefined,
    setInterval: () => 'loading-timer',
    clearInterval: (timer) => {
      clearedTimers.push(timer);
    },
  });

  const prewarmPromise = prewarm({ showLoadingOsd: true });
  shouldShowNotification = false;
  jlptDeferred.resolve();
  freqDeferred.resolve();
  await prewarmPromise;

  assert.deepEqual(clearedTimers, ['loading-timer']);
});

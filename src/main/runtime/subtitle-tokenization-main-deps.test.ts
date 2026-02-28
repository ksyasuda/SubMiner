import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildTokenizerDepsMainHandler,
  createCreateMecabTokenizerAndCheckMainHandler,
  createPrewarmSubtitleDictionariesMainHandler,
} from './subtitle-tokenization-main-deps';

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

test('tokenizer deps builder records known-word lookups and maps readers', () => {
  const calls: string[] = [];
  const deps = createBuildTokenizerDepsMainHandler({
    getYomitanExt: () => null,
    getYomitanParserWindow: () => null,
    setYomitanParserWindow: () => calls.push('set-window'),
    getYomitanParserReadyPromise: () => null,
    setYomitanParserReadyPromise: () => calls.push('set-ready'),
    getYomitanParserInitPromise: () => null,
    setYomitanParserInitPromise: () => calls.push('set-init'),
    isKnownWord: (text) => text === 'known',
    recordLookup: (hit) => calls.push(`lookup:${hit}`),
    getKnownWordMatchMode: () => 'surface',
    getNPlusOneEnabled: () => true,
    getMinSentenceWordsForNPlusOne: () => 3,
    getJlptLevel: () => 'N2',
    getJlptEnabled: () => true,
    getFrequencyDictionaryEnabled: () => true,
    getFrequencyRank: () => 5,
    getYomitanGroupDebugEnabled: () => false,
    getMecabTokenizer: () => null,
  })();

  assert.equal(deps.isKnownWord('known'), true);
  assert.equal(deps.isKnownWord('unknown'), false);
  deps.setYomitanParserWindow(null);
  deps.setYomitanParserReadyPromise(null);
  deps.setYomitanParserInitPromise(null);
  assert.equal(deps.getNPlusOneEnabled?.(), true);
  assert.equal(deps.getMinSentenceWordsForNPlusOne?.(), 3);
  assert.deepEqual(calls, ['lookup:true', 'lookup:false', 'set-window', 'set-ready', 'set-init']);
});

test('mecab tokenizer check creates tokenizer once and runs availability check', async () => {
  const calls: string[] = [];
  type Tokenizer = { id: string };
  let tokenizer: Tokenizer | null = null;
  const run = createCreateMecabTokenizerAndCheckMainHandler<Tokenizer>({
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
  const prewarm = createPrewarmSubtitleDictionariesMainHandler({
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

  const prewarm = createPrewarmSubtitleDictionariesMainHandler({
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
  assert.deepEqual(osdMessages, ['Loading subtitle annotations |', 'Loading subtitle annotations /']);

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

  const prewarm = createPrewarmSubtitleDictionariesMainHandler({
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

test('dictionary prewarm does not show OSD when notifications are disabled', async () => {
  const osdMessages: string[] = [];

  const prewarm = createPrewarmSubtitleDictionariesMainHandler({
    ensureJlptDictionaryLookup: async () => undefined,
    ensureFrequencyDictionaryLookup: async () => undefined,
    shouldShowOsdNotification: () => false,
    showMpvOsd: (message) => {
      osdMessages.push(message);
    },
  });

  await prewarm({ showLoadingOsd: true });

  assert.deepEqual(osdMessages, []);
});

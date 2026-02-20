import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildTokenizerDepsMainHandler,
  createCreateMecabTokenizerAndCheckMainHandler,
  createPrewarmSubtitleDictionariesMainHandler,
} from './subtitle-tokenization-main-deps';

test('tokenizer deps builder records known-word lookups and maps readers', () => {
  const calls: string[] = [];
  const deps = createBuildTokenizerDepsMainHandler({
    getYomitanExt: () => ({ id: 'ext' }),
    getYomitanParserWindow: () => ({ id: 'window' }),
    setYomitanParserWindow: () => calls.push('set-window'),
    getYomitanParserReadyPromise: () => null,
    setYomitanParserReadyPromise: () => calls.push('set-ready'),
    getYomitanParserInitPromise: () => null,
    setYomitanParserInitPromise: () => calls.push('set-init'),
    isKnownWord: (text) => text === 'known',
    recordLookup: (hit) => calls.push(`lookup:${hit}`),
    getKnownWordMatchMode: () => 'exact',
    getMinSentenceWordsForNPlusOne: () => 3,
    getJlptLevel: () => 'N2',
    getJlptEnabled: () => true,
    getFrequencyDictionaryEnabled: () => true,
    getFrequencyRank: () => 5,
    getYomitanGroupDebugEnabled: () => false,
    getMecabTokenizer: () => ({ id: 'mecab' }),
  })();

  assert.equal(deps.isKnownWord('known'), true);
  assert.equal(deps.isKnownWord('unknown'), false);
  deps.setYomitanParserWindow({});
  deps.setYomitanParserReadyPromise(null);
  deps.setYomitanParserInitPromise(null);
  assert.equal(deps.getMinSentenceWordsForNPlusOne(), 3);
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

import assert from 'node:assert/strict';
import test from 'node:test';

// @ts-expect-error Vendored Yomitan translator has no local TypeScript declarations.
import { Translator } from '../vendor/yomitan/js/language/translator.js';

type SortableTermEntry = {
  matchPrimaryReading: boolean;
  maxOriginalTextLength: number;
  textProcessorRuleChainCandidates: unknown[];
  inflectionRuleChainCandidates: unknown[];
  sourceTermExactMatchCount: number;
  frequencyOrder: number;
  dictionaryIndex: number;
  score: number;
  dictionaryAlias: string;
  headwords: Array<{ term: string }>;
  definitions: Array<{ dictionary: string }>;
};

type SortableDefinition = {
  dictionary: string;
  dictionaryAlias: string;
  frequencyOrder: number;
  dictionaryIndex: number;
  score: number;
  headwordIndices: number[];
  index: number;
};

test('Translator prioritizes SubMiner term entries without changing dictionary index order', () => {
  const translator = new Translator({});
  const entries: SortableTermEntry[] = [
    {
      matchPrimaryReading: true,
      maxOriginalTextLength: 4,
      textProcessorRuleChainCandidates: [],
      inflectionRuleChainCandidates: [],
      sourceTermExactMatchCount: 1,
      frequencyOrder: 0,
      dictionaryIndex: 0,
      score: 10,
      dictionaryAlias: 'JMdict',
      headwords: [{ term: 'アイリス' }],
      definitions: [{ dictionary: 'JMdict' }],
    },
    {
      matchPrimaryReading: true,
      maxOriginalTextLength: 4,
      textProcessorRuleChainCandidates: [],
      inflectionRuleChainCandidates: [],
      sourceTermExactMatchCount: 1,
      frequencyOrder: 99,
      dictionaryIndex: 99,
      score: 1,
      dictionaryAlias: 'SubMiner Character Dictionary',
      headwords: [{ term: 'アイリス' }],
      definitions: [{ dictionary: 'SubMiner Character Dictionary' }],
    },
  ];

  translator._sortTermDictionaryEntries(entries as unknown[]);

  assert.equal(entries[0]?.dictionaryAlias, 'SubMiner Character Dictionary');
});

test('Translator prioritizes SubMiner definitions without changing dictionary index order', () => {
  const translator = new Translator({});
  const definitions: SortableDefinition[] = [
    {
      dictionary: 'JMdict',
      dictionaryAlias: 'JMdict',
      frequencyOrder: 0,
      dictionaryIndex: 0,
      score: 10,
      headwordIndices: [0],
      index: 0,
    },
    {
      dictionary: 'SubMiner Character Dictionary',
      dictionaryAlias: 'SubMiner Character Dictionary',
      frequencyOrder: 99,
      dictionaryIndex: 99,
      score: 1,
      headwordIndices: [0],
      index: 1,
    },
  ];

  translator._sortTermDictionaryEntryDefinitions(definitions as unknown[]);

  assert.equal(definitions[0]?.dictionaryAlias, 'SubMiner Character Dictionary');
});

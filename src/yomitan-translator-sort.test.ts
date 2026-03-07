import assert from 'node:assert/strict';
import path from 'node:path';
import test from 'node:test';
import { pathToFileURL } from 'node:url';

import { resolveYomitanExtensionPath } from './core/services/yomitan-extension-paths';

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

async function loadTranslator(): Promise<{ new (...args: unknown[]): { [key: string]: unknown } }> {
  const yomitanRoot = resolveYomitanExtensionPath({ cwd: process.cwd() });
  assert.ok(yomitanRoot, 'Run `bun run build:yomitan` before Yomitan integration tests.');
  const module = await import(
    pathToFileURL(path.join(yomitanRoot, 'js', 'language', 'translator.js')).href
  );
  return module.Translator as { new (...args: unknown[]): { [key: string]: unknown } };
}

test('Translator prioritizes SubMiner term entries without changing dictionary index order', async () => {
  const Translator = await loadTranslator();
  const translator = new Translator({}) as {
    _sortTermDictionaryEntries: (entries: unknown[]) => void;
  };
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

test('Translator prioritizes SubMiner definitions without changing dictionary index order', async () => {
  const Translator = await loadTranslator();
  const translator = new Translator({}) as {
    _sortTermDictionaryEntryDefinitions: (definitions: unknown[]) => void;
  };
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

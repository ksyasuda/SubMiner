import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';

import { DEFAULT_CONFIG } from './config';
import { RuntimeOptionsManager } from './runtime-options';

test('SM-012 runtime options path does not use JSON serialize-clone helpers', () => {
  const source = fs.readFileSync(path.join(process.cwd(), 'src/runtime-options.ts'), 'utf-8');
  assert.equal(source.includes('JSON.parse(JSON.stringify('), false);
});

test('RuntimeOptionsManager returns detached effective Anki config copies', () => {
  const baseConfig = {
    deck: 'Mining',
    note: 'Sentence',
    tags: ['SubMiner'],
    behavior: {
      autoUpdateNewCards: true,
      updateIntervalMs: 5000,
    },
    fieldMapping: {
      sentence: 'Sentence',
      meaning: 'Meaning',
      audio: 'Audio',
      image: 'Image',
      context: 'Context',
      source: 'Source',
      definition: 'Definition',
      sequence: 'Sequence',
      contextSecondary: 'ContextSecondary',
      contextTertiary: 'ContextTertiary',
      primarySpelling: 'PrimarySpelling',
      primaryReading: 'PrimaryReading',
      wordSpelling: 'WordSpelling',
      wordReading: 'WordReading',
    },
    duplicates: {
      mode: 'note' as const,
      scope: 'deck' as const,
      allowedFields: [],
    },
    ai: {
      enabled: false,
      model: '',
      systemPrompt: '',
    },
  };

  const manager = new RuntimeOptionsManager(() => structuredClone(baseConfig), {
    applyAnkiPatch: () => undefined,
    onOptionsChanged: () => undefined,
  });

  const effective = manager.getEffectiveAnkiConnectConfig();
  effective.tags!.push('mutated');
  effective.behavior!.autoUpdateNewCards = false;

  assert.deepEqual(baseConfig.tags, ['SubMiner']);
  assert.equal(baseConfig.behavior.autoUpdateNewCards, true);
});

test('RuntimeOptionsManager keeps known-word and n+1 annotation toggles separate', () => {
  const baseConfig = structuredClone(DEFAULT_CONFIG.ankiConnect);
  const patches: unknown[] = [];
  const manager = new RuntimeOptionsManager(() => structuredClone(baseConfig), {
    applyAnkiPatch: (patch) => {
      patches.push(patch);
    },
    onOptionsChanged: () => undefined,
  });

  assert.equal(
    manager.setOptionValue('subtitle.annotation.knownWords.highlightEnabled', true).ok,
    true,
  );
  assert.equal(manager.setOptionValue('subtitle.annotation.nPlusOne', true).ok, true);

  const effective = manager.getEffectiveAnkiConnectConfig();
  assert.equal(effective.knownWords?.highlightEnabled, true);
  assert.equal(effective.nPlusOne?.enabled, true);
  assert.deepEqual(patches, []);
});

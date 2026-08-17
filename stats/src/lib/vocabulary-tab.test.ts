import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const VOCABULARY_TAB_PATH = fileURLToPath(
  new URL('../components/vocabulary/VocabularyTab.tsx', import.meta.url),
);
const VOCABULARY_HOOK_PATH = fileURLToPath(new URL('../hooks/useVocabulary.ts', import.meta.url));

test('VocabularyTab declares all hooks before loading and error early returns', () => {
  const source = fs.readFileSync(VOCABULARY_TAB_PATH, 'utf8');
  const loadingGuardIndex = source.indexOf('if (loading) {');

  assert.notEqual(loadingGuardIndex, -1, 'expected loading early return');

  const hooksAfterLoadingGuard = source
    .slice(loadingGuardIndex)
    .match(/\buse(?:State|Effect|Memo|Callback|Ref|Reducer)\s*\(/g);

  assert.deepEqual(hooksAfterLoadingGuard ?? [], []);
});

test('VocabularyTab uses uncapped server-side data for its charts and card totals', () => {
  const source = fs.readFileSync(VOCABULARY_TAB_PATH, 'utf8');

  assert.match(source, /\} = useVocabulary\(\);/);
  assert.match(source, /charts\?\.topWordsWithoutNames/);
  assert.match(source, /charts\?\.newWordsTimelineWithoutNames/);
  assert.doesNotMatch(source, /buildVocabularySummary\(/);
  assert.match(source, /uniqueWords: summary\?\.uniqueWordsWithoutNames \?\? 0/);
  assert.match(source, /uniqueWords: summary\?\.uniqueWords \?\? 0/);
  assert.match(source, /value=\{summary \? formatNumber\(summary\.uniqueKanji\) : '…'\}/);
});

test('VocabularyTab surfaces aggregate failures with a retry control', () => {
  const source = fs.readFileSync(VOCABULARY_TAB_PATH, 'utf8');

  assert.match(source, /aggregatesError/);
  assert.match(source, /onClick=\{refreshAggregates\}/);
});

test('useVocabulary loads exact card totals without holding up the vocabulary tables', () => {
  const source = fs.readFileSync(VOCABULARY_HOOK_PATH, 'utf8');

  assert.match(
    source,
    /Promise\.allSettled\(\[\s*client\.getVocabulary\(500\),\s*client\.getKanji\(200\),\s*client\.getKnownWords\(\),?\s*\]\)/,
  );
  assert.match(source, /client\s*\.getVocabularySummary\(\)\s*\.then\(/);
  assert.match(source, /client\s*\.getVocabularyCharts\(\)/);
});

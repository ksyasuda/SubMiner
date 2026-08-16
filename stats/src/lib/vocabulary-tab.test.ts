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

test('VocabularyTab uses database-wide summary totals for its stat cards', () => {
  const source = fs.readFileSync(VOCABULARY_TAB_PATH, 'utf8');

  assert.match(
    source,
    /const chartSummary = useMemo\([\s\S]*buildVocabularySummary\(filteredWords, kanji\)[\s\S]*\[filteredWords, kanji\][\s\S]*\);/,
  );
  assert.match(
    source,
    /const \{ words, kanji, knownWords, summary, loading, error, reload \} = useVocabulary\(\);/,
  );
  assert.match(source, /uniqueWords: summary\?\.uniqueWordsWithoutNames \?\? 0/);
  assert.match(source, /uniqueWords: summary\?\.uniqueWords \?\? 0/);
  assert.match(source, /value=\{summary \? formatNumber\(summary\.uniqueKanji\) : '…'\}/);
});

test('useVocabulary loads exact card totals without holding up the vocabulary tables', () => {
  const source = fs.readFileSync(VOCABULARY_HOOK_PATH, 'utf8');

  assert.match(
    source,
    /Promise\.allSettled\(\[\s*client\.getVocabulary\(500\),\s*client\.getKanji\(200\),\s*client\.getKnownWords\(\),?\s*\]\)/,
  );
  assert.match(source, /void client\s*\.getVocabularySummary\(\)\s*\.then\(/);
});

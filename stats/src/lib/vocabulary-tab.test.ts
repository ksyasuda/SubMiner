import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const VOCABULARY_TAB_PATH = fileURLToPath(
  new URL('../components/vocabulary/VocabularyTab.tsx', import.meta.url),
);

test('VocabularyTab declares all hooks before loading and error early returns', () => {
  const source = fs.readFileSync(VOCABULARY_TAB_PATH, 'utf8');
  const loadingGuardIndex = source.indexOf('if (loading) {');

  assert.notEqual(loadingGuardIndex, -1, 'expected loading early return');

  const hooksAfterLoadingGuard = source
    .slice(loadingGuardIndex)
    .match(/\buse(?:State|Effect|Memo|Callback|Ref|Reducer)\s*\(/g);

  assert.deepEqual(hooksAfterLoadingGuard ?? [], []);
});

test('VocabularyTab memoizes summary and known-word aggregate calculations', () => {
  const source = fs.readFileSync(VOCABULARY_TAB_PATH, 'utf8');

  assert.match(
    source,
    /const summary = useMemo\([\s\S]*buildVocabularySummary\(filteredWords, kanji\)[\s\S]*\[filteredWords, kanji\][\s\S]*\);/,
  );
  assert.match(
    source,
    /const knownWordCount = useMemo\(\(\) => \{[\s\S]*for \(const w of filteredWords\) \{[\s\S]*knownWords\.has\(w\.headword\)[\s\S]*\}\s*return count;\s*\}, \[filteredWords, knownWords\]\);/,
  );
});

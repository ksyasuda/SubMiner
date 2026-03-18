import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';

const VOCABULARY_TAB_PATH = path.resolve(
  import.meta.dir,
  '../components/vocabulary/VocabularyTab.tsx',
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

import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import { formatSentenceSearchMatchCountLabel } from './SearchTab';

const SEARCH_TAB_PATH = path.resolve(path.dirname(fileURLToPath(import.meta.url)), 'SearchTab.tsx');

test('formatSentenceSearchMatchCountLabel uses singular label for one result', () => {
  assert.equal(formatSentenceSearchMatchCountLabel(1), 'Match');
  assert.equal(formatSentenceSearchMatchCountLabel(0), 'Matches');
  assert.equal(formatSentenceSearchMatchCountLabel(2), 'Matches');
});

test('SearchTab forwards stored secondary subtitle text when mining from search results', () => {
  const source = fs.readFileSync(SEARCH_TAB_PATH, 'utf8');

  assert.match(source, /secondaryText:\s*result\.secondaryText/);
});

test('SearchTab enables headword sentence search by default and forwards the toggle', () => {
  const source = fs.readFileSync(SEARCH_TAB_PATH, 'utf8');

  assert.match(
    source,
    /const \[searchByHeadword,\s*setSearchByHeadword\] = useState\(true\);/,
  );
  assert.match(source, /apiClient\s*\.\s*searchSentences\(trimmed,\s*SEARCH_LIMIT,\s*searchByHeadword\)/);
  assert.match(source, /checked=\{searchByHeadword\}/);
  assert.match(source, /setSearchByHeadword\(event\.target\.checked\)/);
});

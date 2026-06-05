import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const SEARCH_TAB_PATH = path.resolve(path.dirname(fileURLToPath(import.meta.url)), 'SearchTab.tsx');

test('SearchTab forwards stored secondary subtitle text when mining from search results', () => {
  const source = fs.readFileSync(SEARCH_TAB_PATH, 'utf8');

  assert.match(source, /secondaryText:\s*result\.secondaryText/);
});

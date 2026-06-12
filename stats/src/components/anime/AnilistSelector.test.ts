import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const ANILIST_SELECTOR_PATH = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  'AnilistSelector.tsx',
);

test('AnilistSelector resyncs query and search state when the initial anime changes', () => {
  const source = fs.readFileSync(ANILIST_SELECTOR_PATH, 'utf8');

  assert.match(source, /setQuery\(normalizedInitialQuery\)/);
  assert.match(source, /setResults\(\[\]\)/);
  assert.match(source, /setLoading\(false\)/);
  assert.match(source, /setLinking\(null\)/);
  assert.match(source, /}, \[initialQuery, animeId\]\);/);
});

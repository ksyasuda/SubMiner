import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const KANJI_DETAIL_PANEL_PATH = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  'KanjiDetailPanel.tsx',
);

test('KanjiDetailPanel uses the centered detail modal layout', () => {
  const source = fs.readFileSync(KANJI_DETAIL_PANEL_PATH, 'utf8');

  assert.match(source, /fixed inset-0 z-40 flex items-center justify-center p-4/);
  assert.match(source, /relative flex max-h-\[85vh\] w-full max-w-2xl flex-col/);
  assert.doesNotMatch(source, /absolute right-0 top-0 h-full/);
});

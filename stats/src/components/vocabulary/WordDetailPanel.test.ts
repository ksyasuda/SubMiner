import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const WORD_DETAIL_PANEL_PATH = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  'WordDetailPanel.tsx',
);

test('WordDetailPanel uses the shared stats mining payload builder', () => {
  const source = fs.readFileSync(WORD_DETAIL_PANEL_PATH, 'utf8');

  assert.match(source, /buildStatsMineCardParams/);
  assert.match(source, /getStatsMineCardUnavailableReason/);
  assert.match(source, /buildStatsMineCardParams\(\s*occ,\s*data!\.detail\.headword,\s*mode\s*\)/);
});

test('WordDetailPanel shows partial media mining errors instead of silent success', () => {
  const source = fs.readFileSync(WORD_DETAIL_PANEL_PATH, 'utf8');

  assert.match(source, /getStatsMineCardError/);
  assert.match(source, /const responseError = getStatsMineCardError\(result\);/);
});

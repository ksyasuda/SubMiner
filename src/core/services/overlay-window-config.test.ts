import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

test('overlay window config explicitly disables renderer sandbox for preload compatibility', () => {
  const sourcePath = path.join(process.cwd(), 'src/core/services/overlay-window.ts');
  const source = fs.readFileSync(sourcePath, 'utf8');

  assert.match(source, /webPreferences:\s*\{[\s\S]*sandbox:\s*false[\s\S]*\}/m);
});

test('overlay window config uses the provided Yomitan session when available', () => {
  const sourcePath = path.join(process.cwd(), 'src/core/services/overlay-window.ts');
  const source = fs.readFileSync(sourcePath, 'utf8');

  assert.match(source, /session:\s*options\.yomitanSession\s*\?\?\s*undefined/);
});

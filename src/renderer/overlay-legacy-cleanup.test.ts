import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

function readWorkspaceFile(relativePath: string): string {
  return fs.readFileSync(path.join(process.cwd(), relativePath), 'utf8');
}

test('keyboard chord map no longer emits legacy invisible overlay script messages', () => {
  const keyboardSource = readWorkspaceFile('src/renderer/handlers/keyboard.ts');
  assert.doesNotMatch(keyboardSource, /subminer-toggle-invisible/);
  assert.doesNotMatch(keyboardSource, /subminer-show-invisible/);
  assert.doesNotMatch(keyboardSource, /subminer-hide-invisible/);
});

test('overlay layer contracts no longer advertise invisible renderer layer', () => {
  const typesSource = readWorkspaceFile('src/types.ts');
  assert.doesNotMatch(typesSource, /export type OverlayLayer = 'visible' \| 'invisible'/);
  assert.doesNotMatch(
    typesSource,
    /getOverlayLayer:\s*\(\)\s*=>\s*'visible'\s*\|\s*'invisible'\s*\|\s*'modal'\s*\|\s*null/,
  );
});

test('renderer stylesheet no longer contains invisible-layer selectors', () => {
  const cssSource = readWorkspaceFile('src/renderer/style.css');
  assert.doesNotMatch(cssSource, /body\.layer-invisible/);
});

test('renderer stylesheet hides focus chrome on top-level overlay focus targets', () => {
  const cssSource = readWorkspaceFile('src/renderer/style.css');
  assert.match(cssSource, /html:focus,\s*body:focus,\s*#overlay:focus\s*\{[^}]*outline:\s*none;/s);
});

test('top-level readme avoids stale overlay-layers wording', () => {
  const readmeSource = readWorkspaceFile('README.md');
  assert.doesNotMatch(readmeSource, /overlay layers/i);
});

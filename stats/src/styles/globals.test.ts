import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const css = readFileSync(fileURLToPath(new URL('./globals.css', import.meta.url)), 'utf8');

test('stats overlay mode paints an opaque full-viewport background', () => {
  assert.match(css, /html,\s*body,\s*#root\s*\{[^}]*height:\s*100%;/s);
  assert.match(css, /body\.overlay-mode\s*\{[^}]*background-color:\s*var\(--color-ctp-base\);/s);
  assert.doesNotMatch(css, /body\.overlay-mode\s*\{[^}]*rgba\(/s);
  assert.match(
    css,
    /body\.overlay-mode #root\s*\{[^}]*background-color:\s*var\(--color-ctp-base\);/s,
  );
});

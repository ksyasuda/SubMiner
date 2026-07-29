import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';
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

test('reduced motion stops the looping delete indicator rather than speeding it up', () => {
  const reducedMotion = /@media \(prefers-reduced-motion: reduce\)\s*\{(?<body>.*)\n\}/s.exec(css)
    ?.groups?.body;
  assert.ok(reducedMotion, 'expected a prefers-reduced-motion block');

  // Shortening an infinite animation makes it run faster, so the looping rules
  // have to switch it off outright and hold a static state instead.
  assert.match(reducedMotion, /\.animate-indeterminate\s*\{[^}]*animation:\s*none;/s);
  assert.match(reducedMotion, /\.animate-indeterminate\s*\{[^}]*width:\s*100%;/s);
  assert.match(reducedMotion, /\.animate-spin\s*\{[^}]*animation:\s*none;/s);
  assert.doesNotMatch(reducedMotion, /\.animate-indeterminate\s*\{[^}]*animation-duration:/s);
  assert.doesNotMatch(reducedMotion, /\.animate-spin\s*\{[^}]*animation-duration:/s);
});

test('the looping delete indicator keeps its animation for everyone else', () => {
  const beforeMediaQuery = css.slice(0, css.indexOf('@media (prefers-reduced-motion'));
  assert.match(
    beforeMediaQuery,
    /\.animate-indeterminate\s*\{\s*animation:\s*indeterminate-sweep/s,
  );
});

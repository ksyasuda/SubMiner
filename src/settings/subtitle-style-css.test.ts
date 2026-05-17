import test from 'node:test';
import assert from 'node:assert/strict';

import {
  getSubtitleCssManagedConfigPaths,
  parseSubtitleCssDeclarations,
  serializeSubtitleCssDeclarations,
} from './subtitle-style-css';

test('serializeSubtitleCssDeclarations builds primary CSS from config minus default colors', () => {
  const css = serializeSubtitleCssDeclarations('primary', {
    'subtitleStyle.fontFamily': 'M PLUS 1, sans-serif',
    'subtitleStyle.fontSize': 35,
    'subtitleStyle.fontColor': '#cad3f5',
    'subtitleStyle.backgroundColor': 'transparent',
    'subtitleStyle.textShadow': '0 2px 6px rgba(0,0,0,0.9)',
    'subtitleStyle.css': {
      filter: 'drop-shadow(0 0 8px #000)',
      '--subtitle-outline': '1px',
    },
  });

  assert.match(css, /font-family: M PLUS 1, sans-serif;/);
  assert.match(css, /font-size: 35px;/);
  assert.match(css, /text-shadow: 0 2px 6px rgba\(0,0,0,0.9\);/);
  assert.match(css, /filter: drop-shadow\(0 0 8px #000\);/);
  assert.match(css, /--subtitle-outline: 1px;/);
  assert.doesNotMatch(css, /^color:/m);
  assert.doesNotMatch(css, /^background-color:/m);
});

test('serializeSubtitleCssDeclarations builds secondary CSS from secondary config paths', () => {
  const css = serializeSubtitleCssDeclarations('secondary', {
    'subtitleStyle.secondary.fontFamily': 'Noto Sans, sans-serif',
    'subtitleStyle.secondary.fontSize': 24,
    'subtitleStyle.secondary.fontColor': '#cad3f5',
    'subtitleStyle.secondary.backgroundColor': 'transparent',
    'subtitleStyle.secondary.css': {
      'text-transform': 'uppercase',
    },
  });

  assert.match(css, /font-family: Noto Sans, sans-serif;/);
  assert.match(css, /font-size: 24px;/);
  assert.match(css, /text-transform: uppercase;/);
  assert.doesNotMatch(css, /^color:/m);
  assert.doesNotMatch(css, /^background-color:/m);
});

test('parseSubtitleCssDeclarations accepts arbitrary declaration properties', () => {
  assert.deepEqual(
    parseSubtitleCssDeclarations(`
      font-size: 40px;
      text-wrap: balance;
      -webkit-text-stroke: 1px black;
      --subtitle-outline: 1px;
    `),
    {
      ok: true,
      declarations: {
        'font-size': '40px',
        'text-wrap': 'balance',
        '-webkit-text-stroke': '1px black',
        '--subtitle-outline': '1px',
      },
    },
  );
});

test('parseSubtitleCssDeclarations rejects selectors and malformed declarations', () => {
  assert.equal(parseSubtitleCssDeclarations('#subtitleRoot { font-size: 40px; }').ok, false);
  assert.equal(parseSubtitleCssDeclarations('font-size 40px;').ok, false);
});

test('getSubtitleCssManagedConfigPaths excludes color controls', () => {
  assert.ok(getSubtitleCssManagedConfigPaths('primary').includes('subtitleStyle.fontSize'));
  assert.ok(
    getSubtitleCssManagedConfigPaths('secondary').includes('subtitleStyle.secondary.fontSize'),
  );
  assert.equal(
    getSubtitleCssManagedConfigPaths('primary').includes('subtitleStyle.fontColor'),
    false,
  );
  assert.equal(
    getSubtitleCssManagedConfigPaths('secondary').includes('subtitleStyle.secondary.fontColor'),
    false,
  );
});

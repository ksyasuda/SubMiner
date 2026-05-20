import test from 'node:test';
import assert from 'node:assert/strict';

import {
  getSubtitleCssManagedConfigPaths,
  parseSubtitleCssDeclarations,
  serializeSubtitleCssDeclarations,
} from './subtitle-style-css';

test('serializeSubtitleCssDeclarations builds primary CSS from all managed appearance config', () => {
  const css = serializeSubtitleCssDeclarations('primary', {
    'subtitleStyle.fontFamily': 'M PLUS 1, sans-serif',
    'subtitleStyle.fontSize': 35,
    'subtitleStyle.fontColor': '#cad3f5',
    'subtitleStyle.backgroundColor': 'transparent',
    'subtitleStyle.hoverTokenColor': '#f4dbd6',
    'subtitleStyle.hoverTokenBackgroundColor': 'transparent',
    'subtitleStyle.paintOrder': 'stroke fill',
    'subtitleStyle.WebkitTextStroke': '1.5px #000',
    'subtitleStyle.textShadow': '0 2px 6px rgba(0,0,0,0.9)',
    'subtitleStyle.css': {
      filter: 'drop-shadow(0 0 8px #000)',
      '--subtitle-outline': '1px',
    },
  });

  assert.match(css, /font-family: M PLUS 1, sans-serif;/);
  assert.match(css, /font-size: 35px;/);
  assert.match(css, /color: #cad3f5;/);
  assert.match(css, /background-color: transparent;/);
  assert.match(css, /--subtitle-hover-token-color: #f4dbd6;/);
  assert.match(css, /--subtitle-hover-token-background-color: transparent;/);
  assert.match(css, /paint-order: stroke fill;/);
  assert.match(css, /-webkit-text-stroke: 1.5px #000;/);
  assert.doesNotMatch(css, /--subtitle-known-word-color:/);
  assert.doesNotMatch(css, /--subtitle-n-plus-one-color:/);
  assert.doesNotMatch(css, /--subtitle-name-match-color:/);
  assert.doesNotMatch(css, /--subtitle-jlpt-n1-color:/);
  assert.doesNotMatch(css, /--subtitle-frequency-single-color:/);
  assert.doesNotMatch(css, /--subtitle-frequency-band-1-color:/);
  assert.match(css, /text-shadow: 0 2px 6px rgba\(0,0,0,0.9\);/);
  assert.match(css, /filter: drop-shadow\(0 0 8px #000\);/);
  assert.match(css, /--subtitle-outline: 1px;/);
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
  assert.match(css, /color: #cad3f5;/);
  assert.match(css, /background-color: transparent;/);
  assert.match(css, /text-transform: uppercase;/);
});

test('serializeSubtitleCssDeclarations builds sidebar CSS from subtitle sidebar config paths', () => {
  const css = serializeSubtitleCssDeclarations('sidebar', {
    'subtitleSidebar.fontFamily': 'M PLUS 1, sans-serif',
    'subtitleSidebar.fontSize': 16,
    'subtitleSidebar.textColor': '#cad3f5',
    'subtitleSidebar.backgroundColor': 'rgba(73, 77, 100, 0.9)',
    'subtitleSidebar.opacity': 0.95,
    'subtitleSidebar.maxWidth': 420,
    'subtitleSidebar.timestampColor': '#a5adcb',
    'subtitleSidebar.activeLineColor': '#f5bde6',
    'subtitleSidebar.activeLineBackgroundColor': 'rgba(138, 173, 244, 0.22)',
    'subtitleSidebar.hoverLineBackgroundColor': 'rgba(54, 58, 79, 0.84)',
    'subtitleSidebar.css': {
      'font-size': '18px',
      'text-wrap': 'pretty',
    },
  });

  assert.match(css, /font-family: M PLUS 1, sans-serif;/);
  assert.match(css, /font-size: 18px;/);
  assert.match(css, /color: #cad3f5;/);
  assert.match(css, /background-color: rgba\(73, 77, 100, 0.9\);/);
  assert.match(css, /opacity: 0.95;/);
  assert.match(css, /--subtitle-sidebar-max-width: 420px;/);
  assert.match(css, /--subtitle-sidebar-timestamp-color: #a5adcb;/);
  assert.match(css, /--subtitle-sidebar-active-line-color: #f5bde6;/);
  assert.match(css, /--subtitle-sidebar-active-background-color: rgba\(138, 173, 244, 0.22\);/);
  assert.match(css, /--subtitle-sidebar-hover-background-color: rgba\(54, 58, 79, 0.84\);/);
  assert.match(css, /text-wrap: pretty;/);
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

test('getSubtitleCssManagedConfigPaths includes all CSS-editor-owned appearance controls', () => {
  assert.ok(!getSubtitleCssManagedConfigPaths('primary').includes(''));
  assert.ok(!getSubtitleCssManagedConfigPaths('secondary').includes(''));
  assert.ok(!getSubtitleCssManagedConfigPaths('sidebar').includes(''));
  assert.ok(getSubtitleCssManagedConfigPaths('primary').includes('subtitleStyle.fontSize'));
  assert.ok(
    getSubtitleCssManagedConfigPaths('secondary').includes('subtitleStyle.secondary.fontSize'),
  );
  assert.ok(getSubtitleCssManagedConfigPaths('sidebar').includes('subtitleSidebar.fontSize'));
  assert.ok(getSubtitleCssManagedConfigPaths('sidebar').includes('subtitleSidebar.textColor'));
  assert.ok(getSubtitleCssManagedConfigPaths('sidebar').includes('subtitleSidebar.maxWidth'));
  assert.equal(
    getSubtitleCssManagedConfigPaths('primary').includes('subtitleStyle.fontColor'),
    true,
  );
  assert.equal(
    getSubtitleCssManagedConfigPaths('secondary').includes('subtitleStyle.secondary.fontColor'),
    true,
  );
  assert.equal(
    getSubtitleCssManagedConfigPaths('primary').includes('subtitleStyle.backgroundColor'),
    true,
  );
  assert.equal(
    getSubtitleCssManagedConfigPaths('secondary').includes(
      'subtitleStyle.secondary.backgroundColor',
    ),
    true,
  );
  assert.equal(
    getSubtitleCssManagedConfigPaths('primary').includes('subtitleStyle.hoverTokenColor'),
    true,
  );
  assert.equal(
    getSubtitleCssManagedConfigPaths('primary').includes('subtitleStyle.hoverTokenBackgroundColor'),
    true,
  );
  assert.equal(
    getSubtitleCssManagedConfigPaths('primary').includes('subtitleStyle.paintOrder'),
    true,
  );
  assert.equal(
    getSubtitleCssManagedConfigPaths('primary').includes('subtitleStyle.WebkitTextStroke'),
    true,
  );
  assert.equal(
    getSubtitleCssManagedConfigPaths('primary').includes('subtitleStyle.knownWordColor'),
    false,
  );
  assert.equal(
    getSubtitleCssManagedConfigPaths('primary').includes('subtitleStyle.nPlusOneColor'),
    false,
  );
  assert.equal(
    getSubtitleCssManagedConfigPaths('primary').includes('subtitleStyle.nameMatchColor'),
    false,
  );
  assert.equal(
    getSubtitleCssManagedConfigPaths('primary').includes('subtitleStyle.jlptColors.N1'),
    false,
  );
  assert.equal(
    getSubtitleCssManagedConfigPaths('primary').includes(
      'subtitleStyle.frequencyDictionary.singleColor',
    ),
    false,
  );
  assert.equal(
    getSubtitleCssManagedConfigPaths('primary').includes(
      'subtitleStyle.frequencyDictionary.bandedColors',
    ),
    false,
  );
});

import assert from 'node:assert/strict';
import test from 'node:test';
import { injectTexthookerBootstrapHtml, type TexthookerBootstrapSettings } from './texthooker';

test('injectTexthookerBootstrapHtml injects websocket bootstrap before head close', () => {
  const html = '<html><head><title>Texthooker</title></head><body></body></html>';
  const settings: TexthookerBootstrapSettings = {
    enableKnownWordColoring: true,
    enableNPlusOneColoring: true,
    enableNameMatchColoring: true,
    enableFrequencyColoring: true,
    enableJlptColoring: true,
    characterDictionaryEnabled: true,
    knownWordColor: '#a6da95',
    nPlusOneColor: '#c6a0f6',
    nameMatchColor: '#f5bde6',
    hoverTokenColor: '#f4dbd6',
    hoverTokenBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    jlptColors: {
      N1: '#ed8796',
      N2: '#f5a97f',
      N3: '#f9e2af',
      N4: '#a6e3a1',
      N5: '#8aadf4',
    },
    frequencyDictionary: {
      singleColor: '#f5a97f',
      bandedColors: ['#ed8796', '#f5a97f', '#f9e2af', '#8bd5ca', '#8aadf4'],
    },
  };
  const actual = injectTexthookerBootstrapHtml(html, 'ws://127.0.0.1:6678', settings);

  assert.match(
    actual,
    /window\.localStorage\.setItem\('bannou-texthooker-websocketUrl', "ws:\/\/127\.0\.0\.1:6678"\)/,
  );
  assert.match(
    actual,
    /window\.localStorage\.setItem\('bannou-texthooker-enableKnownWordColoring', "1"\)/,
  );
  assert.match(
    actual,
    /window\.localStorage\.setItem\('bannou-texthooker-enableNPlusOneColoring', "1"\)/,
  );
  assert.match(
    actual,
    /window\.localStorage\.setItem\('bannou-texthooker-enableNameMatchColoring', "1"\)/,
  );
  assert.match(
    actual,
    /window\.localStorage\.setItem\('bannou-texthooker-enableFrequencyColoring', "1"\)/,
  );
  assert.match(
    actual,
    /window\.localStorage\.setItem\('bannou-texthooker-enableJlptColoring', "1"\)/,
  );
  assert.match(
    actual,
    /window\.localStorage\.setItem\('bannou-texthooker-characterDictionaryEnabled', "1"\)/,
  );
  assert.match(actual, /--subminer-known-word-color:\s*#a6da95;/);
  assert.match(actual, /--subminer-n-plus-one-color:\s*#c6a0f6;/);
  assert.match(actual, /--subminer-name-match-color:\s*#f5bde6;/);
  assert.match(actual, /--subminer-jlpt-n1-color:\s*#ed8796;/);
  assert.match(actual, /--subminer-frequency-band-4-color:\s*#8bd5ca;/);
  assert.match(actual, /--sm-token-hover-bg:\s*rgba\(54, 58, 79, 0\.84\);/);
  assert.match(actual, /p \.word\.word-known\s*\{\s*color:\s*var\(--subminer-known-word-color\);/);
  assert.ok(actual.indexOf('</script></head>') !== -1);
  assert.ok(actual.includes('bannou-texthooker-websocketUrl'));
});

test('injectTexthookerBootstrapHtml leaves html unchanged without websocketUrl', () => {
  const html = '<html><head></head><body></body></html>';

  assert.equal(injectTexthookerBootstrapHtml(html), html);
});

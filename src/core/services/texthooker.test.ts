import assert from 'node:assert/strict';
import test from 'node:test';
import { injectTexthookerBootstrapHtml } from './texthooker';

test('injectTexthookerBootstrapHtml injects websocket bootstrap before head close', () => {
  const html = '<html><head><title>Texthooker</title></head><body></body></html>';

  const actual = injectTexthookerBootstrapHtml(html, 'ws://127.0.0.1:6678');

  assert.match(
    actual,
    /window\.localStorage\.setItem\('bannou-texthooker-websocketUrl', "ws:\/\/127\.0\.0\.1:6678"\)/,
  );
  assert.ok(actual.indexOf('</script></head>') !== -1);
  assert.ok(actual.includes("bannou-texthooker-websocketUrl"));
  assert.ok(!actual.includes('bannou-texthooker-enableKnownWordColoring'));
  assert.ok(!actual.includes('bannou-texthooker-enableNPlusOneColoring'));
  assert.ok(!actual.includes('bannou-texthooker-enableNameMatchColoring'));
  assert.ok(!actual.includes('bannou-texthooker-enableFrequencyColoring'));
  assert.ok(!actual.includes('bannou-texthooker-enableJlptColoring'));
});

test('injectTexthookerBootstrapHtml leaves html unchanged without websocketUrl', () => {
  const html = '<html><head></head><body></body></html>';

  assert.equal(injectTexthookerBootstrapHtml(html), html);
});

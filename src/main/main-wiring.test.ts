import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

function readMainSource(): string {
  return fs.readFileSync(path.join(process.cwd(), 'src/main.ts'), 'utf8');
}

test('manual watched session action starts immersion tracker before marking watched', () => {
  const source = readMainSource();
  const actionBlock = source.match(
    /markActiveVideoWatched:\s*async\s*\(\)\s*=>\s*\{(?<body>[\s\S]*?)\}\s*,/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(actionBlock, /ensureImmersionTrackerStarted\(\);/);
  assert.ok(
    actionBlock.indexOf('ensureImmersionTrackerStarted();') <
      actionBlock.indexOf('markActiveVideoWatched()'),
  );
});

test('media path changes clear rendered subtitle state', () => {
  const source = readMainSource();
  const actionBlock = source.match(
    /updateCurrentMediaPath:\s*\(path\)\s*=>\s*\{(?<body>[\s\S]*?)autoplayReadyGate\.invalidatePendingAutoplayReadyFallbacks\(\);/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(actionBlock, /appState\.currentSubText = '';/);
  assert.match(actionBlock, /appState\.currentSubAssText = '';/);
  assert.match(actionBlock, /appState\.currentSubtitleData = null;/);
  assert.match(actionBlock, /appState\.activeParsedSubtitleCues = \[\];/);
  assert.match(actionBlock, /appState\.activeParsedSubtitleSource = null;/);
  assert.match(actionBlock, /lastObservedTimePos = 0;/);
  assert.match(actionBlock, /broadcastToOverlayWindows\('subtitle:set',/);
  assert.match(actionBlock, /subtitleWsService\.broadcast\(/);
  assert.match(actionBlock, /annotationSubtitleWsService\.broadcast\(/);
  assert.ok(
    actionBlock.indexOf('appState.currentSubtitleData = null;') <
      actionBlock.indexOf("broadcastToOverlayWindows('subtitle:set'"),
  );
});

test('main process uses one shared mpv plugin runtime config helper', () => {
  const source = readMainSource();
  assert.match(source, /function getMpvPluginRuntimeConfig\(\)/);
  assert.equal((source.match(/socketPath: appState\.mpvSocketPath/g) ?? []).length, 1);
  assert.equal(
    (source.match(/binaryPath: getResolvedConfig\(\)\.mpv\.subminerBinaryPath/g) ?? []).length,
    0,
  );
});

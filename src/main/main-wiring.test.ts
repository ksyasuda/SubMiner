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
    /markActiveVideoWatched: async \(\) => \{(?<body>[\s\S]*?)\n    \},/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(actionBlock, /ensureImmersionTrackerStarted\(\);/);
  assert.ok(
    actionBlock.indexOf('ensureImmersionTrackerStarted();') <
      actionBlock.indexOf('markActiveVideoWatched()'),
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

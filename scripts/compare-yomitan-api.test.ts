import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';

const source = readFileSync('scripts/compare-yomitan-api.ts', 'utf8');

test('creates the configured Yomitan user-data directory before Electron uses it', () => {
  const setPathIndex = source.indexOf(
    "electronModule.app.setPath('userData', options.yomitanUserDataPath)",
  );
  assert.notEqual(setPathIndex, -1);

  const mkdirIndex = source.lastIndexOf(
    'fs.mkdirSync(options.yomitanUserDataPath, { recursive: true })',
    setPathIndex,
  );
  assert.ok(mkdirIndex > -1, 'must create the user-data directory');
  assert.ok(mkdirIndex < setPathIndex, 'must create the directory before app.setPath');
});

test('lets output drain by using exitCode in the top-level completion handlers', () => {
  const completionHandlers = source.slice(source.lastIndexOf('\nmain()'));
  assert.match(completionHandlers, /process\.exitCode = process\.exitCode \?\? 0/);
  assert.match(completionHandlers, /process\.exitCode = 1/);
  assert.doesNotMatch(completionHandlers, /process\.exit\(/);
});

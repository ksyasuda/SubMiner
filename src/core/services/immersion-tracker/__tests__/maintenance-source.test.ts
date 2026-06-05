import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';

test('getRollupGroupsForSessions uses only localtime rollup keys', () => {
  const source = fs.readFileSync(
    path.join(process.cwd(), 'src/core/services/immersion-tracker/maintenance.ts'),
    'utf8',
  );
  const start = source.indexOf('export function getRollupGroupsForSessions');
  const end = source.indexOf('export function refreshRollupsForGroupsInTransaction');
  const functionSource = source.slice(start, end);

  assert.match(functionSource, /'unixepoch', 'localtime'/);
  assert.doesNotMatch(functionSource, /UNION/);
  assert.doesNotMatch(functionSource, /86400000/);
});

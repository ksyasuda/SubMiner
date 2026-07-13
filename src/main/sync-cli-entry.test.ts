import test from 'node:test';
import assert from 'node:assert/strict';
import { handleSyncCliAtEntry } from './sync-cli';

test('handleSyncCliAtEntry exits through the process exit dependency', async () => {
  const exits: number[] = [];
  const handled = await handleSyncCliAtEntry(
    ['SubMiner', '--sync-cli', 'sync', '--make-temp'],
    {},
    '0.18.0',
    {
      run: async () => 7,
      exit: (code) => {
        exits.push(code);
      },
    },
  );

  assert.equal(handled, true);
  assert.deepEqual(exits, [7]);
});

test('handleSyncCliAtEntry ignores normal GUI startup', async () => {
  const handled = await handleSyncCliAtEntry(['SubMiner', '--start'], {}, '0.18.0', {
    run: async () => {
      throw new Error('should not run');
    },
    exit: () => {
      throw new Error('should not exit');
    },
  });

  assert.equal(handled, false);
});

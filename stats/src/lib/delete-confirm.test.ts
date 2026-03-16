import assert from 'node:assert/strict';
import test from 'node:test';
import { confirmEpisodeDelete, confirmSessionDelete } from './delete-confirm';

test('confirmSessionDelete uses the shared session delete warning copy', () => {
  const calls: string[] = [];
  const originalConfirm = globalThis.confirm;
  globalThis.confirm = ((message?: string) => {
    calls.push(message ?? '');
    return true;
  }) as typeof globalThis.confirm;

  try {
    assert.equal(confirmSessionDelete(), true);
    assert.deepEqual(calls, ['Delete this session and all associated data?']);
  } finally {
    globalThis.confirm = originalConfirm;
  }
});

test('confirmEpisodeDelete includes the episode title in the shared warning copy', () => {
  const calls: string[] = [];
  const originalConfirm = globalThis.confirm;
  globalThis.confirm = ((message?: string) => {
    calls.push(message ?? '');
    return false;
  }) as typeof globalThis.confirm;

  try {
    assert.equal(confirmEpisodeDelete('Episode 4'), false);
    assert.deepEqual(calls, ['Delete "Episode 4" and all its sessions?']);
  } finally {
    globalThis.confirm = originalConfirm;
  }
});

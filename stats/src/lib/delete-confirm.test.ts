import assert from 'node:assert/strict';
import test from 'node:test';
import {
  confirmBucketDelete,
  confirmDayGroupDelete,
  confirmEpisodeDelete,
  confirmSessionDelete,
} from './delete-confirm';

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

test('confirmDayGroupDelete includes the day label and count in the warning copy', () => {
  const calls: string[] = [];
  const originalConfirm = globalThis.confirm;
  globalThis.confirm = ((message?: string) => {
    calls.push(message ?? '');
    return true;
  }) as typeof globalThis.confirm;

  try {
    assert.equal(confirmDayGroupDelete('Today', 3), true);
    assert.deepEqual(calls, ['Delete all 3 sessions from Today and all associated data?']);
  } finally {
    globalThis.confirm = originalConfirm;
  }
});

test('confirmDayGroupDelete uses singular for one session', () => {
  const calls: string[] = [];
  const originalConfirm = globalThis.confirm;
  globalThis.confirm = ((message?: string) => {
    calls.push(message ?? '');
    return true;
  }) as typeof globalThis.confirm;

  try {
    assert.equal(confirmDayGroupDelete('Yesterday', 1), true);
    assert.deepEqual(calls, ['Delete all 1 session from Yesterday and all associated data?']);
  } finally {
    globalThis.confirm = originalConfirm;
  }
});

test('confirmBucketDelete asks about merging multiple sessions of the same episode', () => {
  const calls: string[] = [];
  const originalConfirm = globalThis.confirm;
  globalThis.confirm = ((message?: string) => {
    calls.push(message ?? '');
    return true;
  }) as typeof globalThis.confirm;

  try {
    assert.equal(confirmBucketDelete('My Episode', 3), true);
    assert.deepEqual(calls, [
      'Delete all 3 sessions of "My Episode" from this day and all associated data?',
    ]);
  } finally {
    globalThis.confirm = originalConfirm;
  }
});

test('confirmBucketDelete uses a clean singular form for one session', () => {
  const calls: string[] = [];
  const originalConfirm = globalThis.confirm;
  globalThis.confirm = ((message?: string) => {
    calls.push(message ?? '');
    return false;
  }) as typeof globalThis.confirm;

  try {
    assert.equal(confirmBucketDelete('Solo Episode', 1), false);
    assert.deepEqual(calls, [
      'Delete this session of "Solo Episode" from this day and all associated data?',
    ]);
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

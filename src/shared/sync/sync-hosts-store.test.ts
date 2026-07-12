import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  createDefaultSyncHostsState,
  getSyncHostsPath,
  isValidSyncHost,
  normalizeSyncHostsState,
  readSyncHostsState,
  recordSyncResult,
  removeSyncHost,
  upsertSyncHost,
  writeSyncHostsState,
} from './sync-hosts-store';

function withTempDir(fn: (dir: string) => void): void {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-sync-hosts-test-'));
  try {
    fn(dir);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

test('createDefaultSyncHostsState returns empty v1 state', () => {
  assert.deepEqual(createDefaultSyncHostsState(), {
    version: 1,
    autoSyncIntervalMinutes: 30,
    hosts: [],
  });
});

test('isValidSyncHost accepts ssh destinations and rejects unsafe values', () => {
  assert.equal(isValidSyncHost('desktop'), true);
  assert.equal(isValidSyncHost('user@10.0.0.5'), true);
  assert.equal(isValidSyncHost('my-alias.local'), true);
  assert.equal(isValidSyncHost(''), false);
  assert.equal(isValidSyncHost('   '), false);
  assert.equal(isValidSyncHost('-oProxyCommand=evil'), false);
  assert.equal(isValidSyncHost('host with spaces'), false);
});

test('upsertSyncHost adds a new host with defaults and preserves fields on update', () => {
  const state = createDefaultSyncHostsState();
  const added = upsertSyncHost(state, { host: 'user@laptop' }, 1000);
  assert.deepEqual(added.hosts, [
    {
      host: 'user@laptop',
      label: null,
      direction: 'both',
      autoSync: false,
      addedAtMs: 1000,
      lastSyncAtMs: null,
      lastSyncStatus: null,
      lastSyncDetail: null,
    },
  ]);
  // original state untouched
  assert.equal(state.hosts.length, 0);

  const synced = recordSyncResult(added, 'user@laptop', {
    atMs: 2000,
    status: 'success',
    detail: 'Sessions merged: 3',
  });
  const updated = upsertSyncHost(synced, { host: 'user@laptop', direction: 'push' }, 3000);
  assert.equal(updated.hosts.length, 1);
  const entry = updated.hosts[0]!;
  assert.equal(entry.direction, 'push');
  assert.equal(entry.addedAtMs, 1000);
  assert.equal(entry.lastSyncAtMs, 2000);
  assert.equal(entry.lastSyncStatus, 'success');
  assert.equal(entry.lastSyncDetail, 'Sessions merged: 3');
});

test('upsertSyncHost trims host and throws on invalid host', () => {
  const state = upsertSyncHost(createDefaultSyncHostsState(), { host: '  desktop  ' }, 1);
  assert.equal(state.hosts[0]!.host, 'desktop');
  assert.throws(() => upsertSyncHost(state, { host: '-bad' }, 2));
});

test('removeSyncHost removes only the matching host', () => {
  let state = createDefaultSyncHostsState();
  state = upsertSyncHost(state, { host: 'a' }, 1);
  state = upsertSyncHost(state, { host: 'b' }, 2);
  const removed = removeSyncHost(state, 'a');
  assert.deepEqual(
    removed.hosts.map((entry) => entry.host),
    ['b'],
  );
  assert.equal(removeSyncHost(removed, 'missing').hosts.length, 1);
});

test('recordSyncResult upserts unknown hosts so CLI syncs are remembered', () => {
  const state = recordSyncResult(createDefaultSyncHostsState(), 'user@server', {
    atMs: 5000,
    status: 'error',
    detail: 'Remote merge failed',
  });
  assert.equal(state.hosts.length, 1);
  const entry = state.hosts[0]!;
  assert.equal(entry.host, 'user@server');
  assert.equal(entry.addedAtMs, 5000);
  assert.equal(entry.lastSyncAtMs, 5000);
  assert.equal(entry.lastSyncStatus, 'error');
  assert.equal(entry.lastSyncDetail, 'Remote merge failed');
});

test('normalizeSyncHostsState drops invalid entries and falls back to defaults', () => {
  assert.deepEqual(normalizeSyncHostsState(null), createDefaultSyncHostsState());
  assert.deepEqual(normalizeSyncHostsState('junk'), createDefaultSyncHostsState());
  assert.deepEqual(normalizeSyncHostsState({ version: 99 }), createDefaultSyncHostsState());

  const normalized = normalizeSyncHostsState({
    version: 1,
    autoSyncIntervalMinutes: 15,
    hosts: [
      {
        host: 'desktop',
        label: 'Desktop PC',
        direction: 'pull',
        autoSync: true,
        addedAtMs: 10,
        lastSyncAtMs: 20,
        lastSyncStatus: 'success',
        lastSyncDetail: 'ok',
      },
      { host: '-bad', direction: 'both', autoSync: false, addedAtMs: 1 },
      { host: 'dupe', direction: 'sideways', autoSync: false, addedAtMs: 1 },
      'not-an-object',
    ],
  });
  assert.equal(normalized.autoSyncIntervalMinutes, 15);
  assert.deepEqual(normalized.hosts, [
    {
      host: 'desktop',
      label: 'Desktop PC',
      direction: 'pull',
      autoSync: true,
      addedAtMs: 10,
      lastSyncAtMs: 20,
      lastSyncStatus: 'success',
      lastSyncDetail: 'ok',
    },
    {
      host: 'dupe',
      label: null,
      direction: 'both',
      autoSync: false,
      addedAtMs: 1,
      lastSyncAtMs: null,
      lastSyncStatus: null,
      lastSyncDetail: null,
    },
  ]);
});

test('read/write round-trips state and tolerates missing or corrupt files', () => {
  withTempDir((root) => {
    const filePath = getSyncHostsPath(root);
    assert.equal(filePath, path.join(root, 'sync-hosts.json'));
    assert.deepEqual(readSyncHostsState(filePath), createDefaultSyncHostsState());

    fs.writeFileSync(filePath, '{corrupt');
    assert.deepEqual(readSyncHostsState(filePath), createDefaultSyncHostsState());

    let state = createDefaultSyncHostsState();
    state = upsertSyncHost(state, { host: 'user@laptop', label: 'Laptop' }, 42);
    writeSyncHostsState(filePath, state);
    assert.deepEqual(readSyncHostsState(filePath), state);
  });
});

test('writeSyncHostsState creates parent directories', () => {
  withTempDir((root) => {
    const filePath = path.join(root, 'nested', 'dir', 'sync-hosts.json');
    writeSyncHostsState(filePath, createDefaultSyncHostsState());
    assert.deepEqual(readSyncHostsState(filePath), createDefaultSyncHostsState());
  });
});

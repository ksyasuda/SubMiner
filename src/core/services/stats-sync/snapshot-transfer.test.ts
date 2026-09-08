import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { randomBytes } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { createSnapshotTransfer } from './snapshot-transfer';
import { createTransferCache, transferCacheKey } from './transfer-cache';

type TransferDeps = NonNullable<Parameters<typeof createSnapshotTransfer>[2]>;

function commandResult(status = 0, stderr = ''): ReturnType<TransferDeps['runRsync']> {
  return { status, stderr, stdout: '', pid: 0, output: [null, '', stderr], signal: null };
}

function makeDeps(overrides: Partial<TransferDeps> = {}): TransferDeps {
  return {
    platform: 'linux',
    runRsync: () => commandResult(),
    runSsh: () => ({ status: 0, stdout: '', stderr: '' }),
    runScp: () => assert.fail('Unexpected scp fallback'),
    ...overrides,
  };
}

test('snapshot transfer falls back when rsync is unavailable or an endpoint is Windows', () => {
  for (const scenario of ['local-missing', 'remote-missing', 'local-windows', 'remote-windows']) {
    const copies: string[][] = [];
    const transfer = createSnapshotTransfer(
      'macbook',
      scenario === 'remote-windows' ? 'windows-cmd' : 'posix',
      makeDeps({
        platform: scenario === 'local-windows' ? 'win32' : 'linux',
        runRsync: () => commandResult(scenario === 'local-missing' ? 1 : 0),
        runSsh: () => ({ status: scenario === 'remote-missing' ? 127 : 0, stdout: '', stderr: '' }),
        runScp: (from, to) => copies.push([from, to]),
      }),
    );
    assert.equal(transfer.kind, 'scp', scenario);
    transfer.copy({
      direction: 'download',
      localPath: '/local.sqlite',
      remotePath: '/remote.sqlite',
    });
    transfer.copy({
      direction: 'upload',
      localPath: '/local.sqlite',
      remotePath: '/remote.sqlite',
    });
    assert.deepEqual(copies, [
      ['macbook:/remote.sqlite', '/local.sqlite'],
      ['/local.sqlite', 'macbook:/remote.sqlite'],
    ]);
  }
});

test('failed rsync transfers report errors without silently retrying through scp', () => {
  const transfer = createSnapshotTransfer(
    'macbook',
    'posix',
    makeDeps({
      runRsync: (args) => commandResult(args.includes('--version') ? 0 : 23, 'Permission denied'),
    }),
  );
  assert.throws(
    () =>
      transfer.copy({
        direction: 'upload',
        localPath: '/local.sqlite',
        remotePath: '/remote.sqlite',
      }),
    /rsync upload failed for macbook: Permission denied/,
  );
  assert.throws(() => createSnapshotTransfer('-oProxyCommand=bad', 'posix', makeDeps()), /option/);
});

const hasRsync = process.platform !== 'win32' && spawnSync('rsync', ['--version']).status === 0;

for (const direction of ['download', 'upload'] as const) {
  test(
    `rsync ${direction} reuses snapshot blocks and preserves the basis`,
    { skip: !hasRsync },
    () => {
      const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-transfer-test-'));
      try {
        const localDir = path.join(dir, 'local');
        const remoteDir = path.join(dir, "remote space ' $(false)");
        fs.mkdirSync(localDir);
        fs.mkdirSync(remoteDir);
        // Emulate SSH's remote shell with real rsync processes, without sshd.
        const remoteShell = path.join(dir, 'remote-shell');
        fs.writeFileSync(remoteShell, '#!/bin/sh\nshift\nexec /bin/sh -c "$*"\n', { mode: 0o700 });
        const localPath = path.join(
          localDir,
          direction === 'download' ? 'incoming' : '',
          'snapshot.sqlite',
        );
        const remotePath = path.join(
          remoteDir,
          direction === 'upload' ? 'incoming' : '',
          'snapshot.sqlite',
        );
        const source = direction === 'download' ? remotePath : localPath;
        const destination = direction === 'download' ? localPath : remotePath;
        const basis = path.join(path.dirname(destination), '..', 'snapshot.sqlite');
        // Incompressible data ensures savings come from matching blocks.
        const original = randomBytes(4 * 1024 * 1024);
        fs.writeFileSync(basis, original);
        fs.mkdirSync(path.dirname(destination), { recursive: true });
        fs.writeFileSync(destination, original);
        const updated = Buffer.from(original);
        updated.fill(42, 65536, 69632);
        fs.writeFileSync(source, updated);
        let stats = '';
        const transfer = createSnapshotTransfer(
          'test-peer',
          'posix',
          makeDeps({
            runRsync: (args) => {
              const result = spawnSync(
                'rsync',
                [
                  `--rsh=${remoteShell}`,
                  ...args.map((arg) => (arg === '--quiet' ? '--stats' : arg)),
                ],
                { encoding: 'utf8', env: { ...process.env, RSYNC_OLD_ARGS: '1', LC_ALL: 'C' } },
              );
              stats = result.stdout;
              return result;
            },
          }),
        );
        assert.equal(transfer.kind, 'rsync');
        transfer.copy({ direction, localPath, remotePath });
        assert.deepEqual(fs.readFileSync(destination), updated);
        assert.deepEqual(fs.readFileSync(basis), original);
        const matched = /Matched data: ([\d,]+) (?:bytes|B)/.exec(stats)?.[1];
        assert.ok(matched, stats);
        assert.ok(Number(matched.replaceAll(',', '')) > original.length * 0.95, stats);

        const coldDir = path.join(dir, 'cold');
        fs.mkdirSync(coldDir);
        const coldDestination = path.join(coldDir, 'incoming', 'snapshot.sqlite');
        transfer.copy({
          direction,
          localPath: direction === 'download' ? coldDestination : localPath,
          remotePath: direction === 'upload' ? coldDestination : remotePath,
        });
        assert.deepEqual(fs.readFileSync(coldDestination), updated);

        // A later sync starts in a new directory and reuses the prior peer's
        // received file even when the source has grown since that transfer.
        const cache = createTransferCache(path.join(dir, 'cache'));
        const key = transferCacheKey('peer');
        cache.remember(key, direction === 'download' ? localDir : remoteDir);
        const nextDir = path.join(dir, 'next');
        cache.seed(key, nextDir);
        const grown = Buffer.concat([updated, randomBytes(4096)]);
        fs.writeFileSync(source, grown);
        const nextDestination = path.join(nextDir, 'incoming', 'snapshot.sqlite');
        transfer.copy({
          direction,
          localPath: direction === 'download' ? nextDestination : localPath,
          remotePath: direction === 'upload' ? nextDestination : remotePath,
        });
        assert.deepEqual(fs.readFileSync(nextDestination), grown);
        const cachedMatches = /Matched data: ([\d,]+) (?:bytes|B)/.exec(stats)?.[1];
        assert.ok(cachedMatches, stats);
        assert.ok(Number(cachedMatches.replaceAll(',', '')) > updated.length * 0.95, stats);

        // A retry must replace stale content even if size and mtime agree.
        updated.fill(43, 131072, 135168);
        fs.writeFileSync(source, updated);
        const timestamp = new Date(1_700_000_000_000);
        fs.utimesSync(source, timestamp, timestamp);
        fs.utimesSync(destination, timestamp, timestamp);
        transfer.copy({ direction, localPath, remotePath });
        assert.deepEqual(fs.readFileSync(destination), updated);
        assert.deepEqual(fs.readFileSync(basis), original);
      } finally {
        fs.rmSync(dir, { recursive: true, force: true });
      }
    },
  );
}

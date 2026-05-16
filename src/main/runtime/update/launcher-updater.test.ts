import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import {
  buildProtectedLauncherUpdateCommand,
  looksLikeSubminerLauncher,
  updateLauncherAtPath,
} from './launcher-updater';

const launcherBytes = Buffer.from('#!/usr/bin/env bash\n# SubMiner launcher\nexec SubMiner "$@"\n');
const launcherHash = createHash('sha256').update(launcherBytes).digest('hex');

test('looksLikeSubminerLauncher rejects unrelated executable content', () => {
  assert.equal(looksLikeSubminerLauncher(Buffer.from('#!/bin/sh\necho nope\n')), false);
  assert.equal(looksLikeSubminerLauncher(Buffer.from('SubMiner launcher binary payload')), true);
});

test('buildProtectedLauncherUpdateCommand quotes sudo curl and chmod paths', () => {
  assert.equal(
    buildProtectedLauncherUpdateCommand(
      "https://github.com/ksyasuda/SubMiner/releases/latest/download/sub miner?sig='abc'",
      "/usr/local/bin/subminer's launcher",
    ),
    "sudo curl -fSL 'https://github.com/ksyasuda/SubMiner/releases/latest/download/sub miner?sig='\\''abc'\\''' -o '/usr/local/bin/subminer'\\''s launcher' && sudo chmod +x '/usr/local/bin/subminer'\\''s launcher'",
  );
});

test('updateLauncherAtPath verifies hash and atomically replaces writable launcher', async () => {
  const writes: Array<{ path: string; data: Buffer }> = [];
  const renames: Array<{ from: string; to: string }> = [];
  const chmods: Array<{ path: string; mode: number }> = [];

  const result = await updateLauncherAtPath({
    launcherPath: '/home/kyle/.local/bin/subminer',
    assetUrl: 'https://example.test/subminer',
    expectedSha256: launcherHash,
    download: async () => launcherBytes,
    fs: {
      readFile: async () => Buffer.from('#!/bin/sh\n# SubMiner launcher\n'),
      stat: async () => ({ isFile: () => true, mode: 0o755 }),
      access: async () => undefined,
      writeFile: async (filePath, data) => {
        writes.push({ path: filePath, data: Buffer.from(data) });
      },
      chmod: async (filePath, mode) => {
        chmods.push({ path: filePath, mode });
      },
      rename: async (from, to) => {
        renames.push({ from, to });
      },
      unlink: async () => undefined,
    },
  });

  assert.equal(result.status, 'updated');
  assert.equal(writes.length, 1);
  assert.equal(writes[0]!.path, '/home/kyle/.local/bin/.subminer.update');
  assert.equal(writes[0]!.data.equals(launcherBytes), true);
  assert.deepEqual(chmods, [{ path: '/home/kyle/.local/bin/.subminer.update', mode: 0o755 }]);
  assert.deepEqual(renames, [
    { from: '/home/kyle/.local/bin/.subminer.update', to: '/home/kyle/.local/bin/subminer' },
  ]);
});

test('updateLauncherAtPath reports protected command without replacing non-writable launcher', async () => {
  const result = await updateLauncherAtPath({
    launcherPath: '/usr/local/bin/subminer',
    assetUrl: 'https://example.test/subminer',
    expectedSha256: launcherHash,
    download: async () => launcherBytes,
    fs: {
      readFile: async () => Buffer.from('#!/bin/sh\n# SubMiner launcher\n'),
      stat: async () => ({ isFile: () => true, mode: 0o755 }),
      access: async () => {
        throw Object.assign(new Error('EACCES'), { code: 'EACCES' });
      },
      writeFile: async () => {
        throw new Error('unexpected write');
      },
      chmod: async () => undefined,
      rename: async () => undefined,
      unlink: async () => undefined,
    },
  });

  assert.equal(result.status, 'protected');
  assert.match(result.command ?? '', /^sudo curl -fSL 'https:\/\/example\.test\/subminer'/);
});

test('updateLauncherAtPath aborts on hash mismatch and suspicious launcher content', async () => {
  const suspicious = await updateLauncherAtPath({
    launcherPath: '/home/kyle/bin/subminer',
    assetUrl: 'https://example.test/subminer',
    expectedSha256: launcherHash,
    download: async () => launcherBytes,
    fs: {
      readFile: async () => Buffer.from('#!/bin/sh\necho not-subminer\n'),
      stat: async () => ({ isFile: () => true, mode: 0o755 }),
      access: async () => undefined,
      writeFile: async () => undefined,
      chmod: async () => undefined,
      rename: async () => undefined,
      unlink: async () => undefined,
    },
  });
  const mismatch = await updateLauncherAtPath({
    launcherPath: '/home/kyle/.local/bin/subminer',
    assetUrl: 'https://example.test/subminer',
    expectedSha256: '0'.repeat(64),
    download: async () => launcherBytes,
    fs: {
      readFile: async () => Buffer.from('#!/bin/sh\n# SubMiner launcher\n'),
      stat: async () => ({ isFile: () => true, mode: 0o755 }),
      access: async () => undefined,
      writeFile: async () => {
        throw new Error('unexpected write');
      },
      chmod: async () => undefined,
      rename: async () => undefined,
      unlink: async () => undefined,
    },
  });

  assert.equal(suspicious.status, 'skipped');
  assert.equal(mismatch.status, 'hash-mismatch');
});

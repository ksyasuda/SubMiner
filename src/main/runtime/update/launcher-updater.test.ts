import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import {
  buildProtectedLauncherUpdateCommand,
  looksLikeSubminerLauncher,
  updateLauncherAtPath,
  updateLauncherFromRelease,
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

test('app-managed wrappers are never overwritten by the standalone release script', async () => {
  const { managedLauncherContent } = await import('../managed-launcher');
  let downloaded = false;
  const result = await updateLauncherAtPath({
    launcherPath: '/home/tester/.local/bin/subminer',
    assetUrl: 'https://example.test/subminer',
    expectedSha256: launcherHash,
    download: async () => {
      downloaded = true;
      return launcherBytes;
    },
    fs: {
      stat: async () => ({ isFile: () => true }),
      readFile: async () =>
        managedLauncherContent({
          platform: 'linux',
          appPath: '/apps/SubMiner.AppImage',
        }),
      access: async () => {
        throw new Error('must not modify wrapper');
      },
      writeFile: async () => {
        throw new Error('must not modify wrapper');
      },
      chmod: async () => {},
      rename: async () => {},
      unlink: async () => {},
    },
  });
  assert.equal(result.status, 'skipped');
  assert.equal(downloaded, false);
});

test('GUI updates defer recognized standalone launcher migration to app startup', async () => {
  let accessed = false;
  let downloaded = false;
  const result = await updateLauncherAtPath({
    launcherPath: '/home/tester/.local/bin/subminer',
    assetUrl: 'https://example.test/subminer',
    expectedSha256: launcherHash,
    deferRecognizedLauncherUpdate: true,
    download: async () => {
      downloaded = true;
      return launcherBytes;
    },
    fs: {
      stat: async () => ({ isFile: () => true }),
      readFile: async () => Buffer.from('#!/bin/sh\n# SubMiner launcher\n'),
      access: async () => {
        accessed = true;
      },
      writeFile: async () => {},
      chmod: async () => {},
      rename: async () => {},
      unlink: async () => {},
    },
  });

  assert.deepEqual(result, {
    status: 'skipped',
    path: '/home/tester/.local/bin/subminer',
    message: 'Launcher migration is deferred until the updated SubMiner app starts.',
  });
  assert.equal(accessed, false);
  assert.equal(downloaded, false);
});

test('release launcher updater propagates GUI migration deferral', async () => {
  let downloaded = false;
  const result = await updateLauncherFromRelease({
    release: {
      tag_name: 'v0.15.0',
      prerelease: false,
      draft: false,
      assets: [{ name: 'subminer', browser_download_url: 'https://example.test/subminer' }],
    },
    sha256Sums: new Map([['subminer', launcherHash]]),
    launcherPath: '/home/tester/.local/bin/subminer',
    deferRecognizedLauncherUpdate: true,
    exists: () => true,
    downloadAsset: async () => {
      downloaded = true;
      return launcherBytes;
    },
    fs: {
      stat: async () => ({ isFile: () => true }),
      readFile: async () => Buffer.from('#!/bin/sh\n# SubMiner launcher\n'),
      access: async () => {
        throw new Error('must not check writability before app startup');
      },
      writeFile: async () => {},
      chmod: async () => {},
      rename: async () => {},
      unlink: async () => {},
    },
  });

  assert.equal(result.status, 'skipped');
  assert.match(result.message ?? '', /deferred until the updated SubMiner app starts/);
  assert.equal(downloaded, false);
});

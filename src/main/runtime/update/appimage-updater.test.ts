import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { updateAppImageFromRelease } from './appimage-updater';

const appImageBytes = Buffer.from('appimage');
const appImageHash = createHash('sha256').update(appImageBytes).digest('hex');

test('updateAppImageFromRelease verifies hash and atomically replaces writable AppImage', async () => {
  const writes: Array<{ path: string; data: Buffer }> = [];
  const chmods: Array<{ path: string; mode: number }> = [];
  const renames: Array<{ from: string; to: string }> = [];

  const result = await updateAppImageFromRelease({
    release: {
      tag_name: 'v0.15.0',
      prerelease: false,
      draft: false,
      assets: [{ name: 'SubMiner.AppImage', browser_download_url: 'https://example.test/app' }],
    },
    sha256Sums: new Map([['SubMiner.AppImage', appImageHash]]),
    appImagePath: '/home/kyle/.local/bin/SubMiner.AppImage',
    downloadAsset: async () => appImageBytes,
    fs: {
      stat: async () => ({
        isFile: () => true,
        mode: 0o755,
      }),
      access: async () => {},
      writeFile: async (targetPath, data) => {
        writes.push({ path: targetPath, data });
      },
      chmod: async (targetPath, mode) => {
        chmods.push({ path: targetPath, mode });
      },
      rename: async (from, to) => {
        renames.push({ from, to });
      },
      unlink: async () => {},
    },
  });

  assert.deepEqual(result, {
    status: 'updated',
    path: '/home/kyle/.local/bin/SubMiner.AppImage',
  });
  assert.deepEqual(writes, [
    { path: '/home/kyle/.local/bin/.SubMiner.AppImage.update', data: appImageBytes },
  ]);
  assert.deepEqual(chmods, [
    { path: '/home/kyle/.local/bin/.SubMiner.AppImage.update', mode: 0o755 },
  ]);
  assert.deepEqual(renames, [
    {
      from: '/home/kyle/.local/bin/.SubMiner.AppImage.update',
      to: '/home/kyle/.local/bin/SubMiner.AppImage',
    },
  ]);
});

test('updateAppImageFromRelease reports protected command without replacing non-writable AppImage', async () => {
  const result = await updateAppImageFromRelease({
    release: {
      tag_name: 'v0.15.0',
      prerelease: false,
      draft: false,
      assets: [{ name: 'SubMiner.AppImage', browser_download_url: 'https://example.test/app' }],
    },
    sha256Sums: new Map([['SubMiner.AppImage', appImageHash]]),
    appImagePath: '/opt/SubMiner/SubMiner.AppImage',
    downloadAsset: async () => appImageBytes,
    fs: {
      stat: async () => ({
        isFile: () => true,
        mode: 0o755,
      }),
      access: async () => {
        throw new Error('EACCES');
      },
      writeFile: async () => {
        throw new Error('unexpected write');
      },
      chmod: async () => {},
      rename: async () => {},
      unlink: async () => {},
    },
  });

  assert.equal(result.status, 'protected');
  assert.equal(result.path, '/opt/SubMiner/SubMiner.AppImage');
  assert.match(result.command ?? '', /^sudo curl -fSL https:\/\/example\.test\/app -o /);
});

test('updateAppImageFromRelease aborts on hash mismatch', async () => {
  const result = await updateAppImageFromRelease({
    release: {
      tag_name: 'v0.15.0',
      prerelease: false,
      draft: false,
      assets: [{ name: 'SubMiner.AppImage', browser_download_url: 'https://example.test/app' }],
    },
    sha256Sums: new Map([['SubMiner.AppImage', '0'.repeat(64)]]),
    appImagePath: '/home/kyle/.local/bin/SubMiner.AppImage',
    downloadAsset: async () => appImageBytes,
    fs: {
      stat: async () => ({
        isFile: () => true,
        mode: 0o755,
      }),
      access: async () => {},
      writeFile: async () => {
        throw new Error('unexpected write');
      },
      chmod: async () => {},
      rename: async () => {},
      unlink: async () => {},
    },
  });

  assert.equal(result.status, 'hash-mismatch');
});

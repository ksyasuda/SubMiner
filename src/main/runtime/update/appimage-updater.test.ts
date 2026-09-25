import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { buildProtectedAppImageUpdateCommand, updateAppImageFromRelease } from './appimage-updater';

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

test('updateAppImageFromRelease reports protected command for a direct non-writable AppImage', async () => {
  const result = await updateAppImageFromRelease({
    release: {
      tag_name: 'v0.15.0',
      prerelease: false,
      draft: false,
      assets: [{ name: 'SubMiner.AppImage', browser_download_url: 'https://example.test/app' }],
    },
    sha256Sums: new Map([['SubMiner.AppImage', appImageHash]]),
    appImagePath: '/usr/local/lib/SubMiner.AppImage',
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
  assert.equal(result.path, '/usr/local/lib/SubMiner.AppImage');
  assert.match(result.command ?? '', /curl -fSL 'https:\/\/example\.test\/app' -o "\$tmp"/);
  assert.match(result.command ?? '', /sha256sum -c -/);
  assert.match(result.command ?? '', /sudo mv "\$tmp" '\/usr\/local\/lib\/SubMiner\.AppImage'/);
});

test('updateAppImageFromRelease leaves canonical and symlinked AUR AppImages to pacman', async () => {
  for (const appImagePath of ['/opt/SubMiner/SubMiner.AppImage', '/usr/bin/SubMiner.AppImage']) {
    let accessed = false;
    const result = await updateAppImageFromRelease({
      release: {
        tag_name: 'v0.15.0',
        prerelease: false,
        draft: false,
        assets: [{ name: 'SubMiner.AppImage', browser_download_url: 'https://example.test/app' }],
      },
      sha256Sums: new Map([['SubMiner.AppImage', appImageHash]]),
      appImagePath,
      downloadAsset: async () => {
        throw new Error('must not download package-managed AppImage');
      },
      fs: {
        realpath: async () => '/opt/SubMiner/SubMiner.AppImage',
        stat: async () => {
          throw new Error('must not stat package-managed AppImage');
        },
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
      path: appImagePath,
      message: 'This AppImage is managed by the subminer-bin system package.',
    });
    assert.equal(accessed, false);
    assert.equal(result.command, undefined);
  }
});

test('buildProtectedAppImageUpdateCommand quotes inputs and verifies checksum before sudo move', () => {
  const command = buildProtectedAppImageUpdateCommand(
    "https://example.test/Sub Miner.AppImage?sig='abc'",
    "/opt/Sub Miner/SubMiner's.AppImage",
    'ABCDEF',
  );

  assert.match(command, /trap 'rm -f "\$tmp"' EXIT/);
  assert.match(
    command,
    /curl -fSL 'https:\/\/example\.test\/Sub Miner\.AppImage\?sig='\\''abc'\\''' -o "\$tmp"/,
  );
  assert.match(command, /printf '%s  %s\\n' 'abcdef' "\$tmp" \| sha256sum -c -/);
  assert.match(command, /sudo mv "\$tmp" '\/opt\/Sub Miner\/SubMiner'\\''s\.AppImage'/);
  assert.match(command, /sudo chmod \+x '\/opt\/Sub Miner\/SubMiner'\\''s\.AppImage'/);
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

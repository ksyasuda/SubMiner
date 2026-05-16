import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  buildProtectedSupportAssetsCommand,
  detectSupportAssetDataDirs,
  updateSupportAssetsFromRelease,
} from './support-assets';

function sha256(data: Buffer): string {
  return createHash('sha256').update(data).digest('hex');
}

function makeSupportAssetsArchive(): { archive: Buffer; tempDir: string } {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-support-assets-test-'));
  fs.mkdirSync(path.join(tempDir, 'assets/themes'), { recursive: true });
  fs.mkdirSync(path.join(tempDir, 'plugin/subminer'), { recursive: true });
  fs.writeFileSync(path.join(tempDir, 'assets/themes/subminer.rasi'), 'new theme\n');
  fs.writeFileSync(path.join(tempDir, 'plugin/subminer/main.lua'), 'new plugin\n');
  execFileSync('tar', ['-czf', 'subminer-assets.tar.gz', 'assets', 'plugin'], { cwd: tempDir });
  return {
    archive: fs.readFileSync(path.join(tempDir, 'subminer-assets.tar.gz')),
    tempDir,
  };
}

test('detectSupportAssetDataDirs only returns Linux rofi theme locations', () => {
  assert.deepEqual(
    detectSupportAssetDataDirs({
      platform: 'darwin',
      homeDir: '/Users/kyle',
    }),
    [],
  );
  assert.deepEqual(
    detectSupportAssetDataDirs({
      platform: 'linux',
      homeDir: '/home/kyle',
      xdgDataHome: '/tmp/xdg-data',
    }),
    ['/tmp/xdg-data/SubMiner', '/usr/local/share/SubMiner', '/usr/share/SubMiner'],
  );
});

test('buildProtectedSupportAssetsCommand cleans up temporary extraction directory', () => {
  const command = buildProtectedSupportAssetsCommand(
    "https://example.test/subminer assets.tar.gz?sig='abc'",
    "/usr/local/share/SubMiner's data",
  );

  assert.match(command, /tmp=\$\(mktemp -d\)/);
  assert.match(command, /trap 'rm -rf "\$tmp"' EXIT/);
  assert.match(
    command,
    /curl -fSL 'https:\/\/example\.test\/subminer assets\.tar\.gz\?sig='\\''abc'\\''' -o "\$tmp\/subminer-assets\.tar\.gz"/,
  );
  assert.match(command, /sudo mkdir -p '\/usr\/local\/share\/SubMiner'\\''s data'\/themes/);
});

test('updateSupportAssetsFromRelease updates only the Linux rofi theme', async () => {
  const xdgDataHome = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-xdg-data-'));
  const dataDir = path.join(xdgDataHome, 'SubMiner');
  fs.mkdirSync(path.join(dataDir, 'themes'), { recursive: true });
  fs.mkdirSync(path.join(dataDir, 'plugin/subminer'), { recursive: true });
  fs.writeFileSync(path.join(dataDir, 'themes/subminer.rasi'), 'old theme\n');
  fs.writeFileSync(path.join(dataDir, 'plugin/subminer/main.lua'), 'old plugin\n');
  const { archive, tempDir } = makeSupportAssetsArchive();

  try {
    const results = await updateSupportAssetsFromRelease({
      release: {
        tag_name: 'v0.15.0',
        assets: [
          {
            name: 'subminer-assets.tar.gz',
            browser_download_url: 'https://example.test/subminer-assets.tar.gz',
          },
        ],
      },
      sha256Sums: new Map([['subminer-assets.tar.gz', sha256(archive)]]),
      downloadAsset: async () => archive,
      platform: 'linux',
      xdgDataHome,
    });

    assert.deepEqual(results, [{ status: 'updated', path: dataDir }]);
    assert.equal(
      fs.readFileSync(path.join(dataDir, 'themes/subminer.rasi'), 'utf8'),
      'new theme\n',
    );
    assert.equal(
      fs.readFileSync(path.join(dataDir, 'plugin/subminer/main.lua'), 'utf8'),
      'old plugin\n',
    );
  } finally {
    fs.rmSync(xdgDataHome, { recursive: true, force: true });
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
});

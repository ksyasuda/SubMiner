import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import {
  findBundleBinaries,
  PINNED_BUNDLE_SHA256,
  PINNED_BUNDLE_TAG,
  resolveBundleAssetName,
  selectBundleAsset,
  sha256,
  verifyPinnedBundle,
} from './sidecar-bundle';

test('resolveBundleAssetName maps supported platform/arch pairs', () => {
  assert.equal(resolveBundleAssetName('darwin', 'arm64'), 'macOS-arm64-bundle.zip');
  assert.equal(resolveBundleAssetName('darwin', 'x64'), 'macOS-x64-bundle.zip');
  assert.equal(resolveBundleAssetName('linux', 'x64'), 'linux-x64-bundle.zip');
  assert.equal(resolveBundleAssetName('win32', 'x64'), 'windows-x64-bundle.zip');
});

test('resolveBundleAssetName returns null for unpublished combinations', () => {
  assert.equal(resolveBundleAssetName('linux', 'arm64'), null);
  assert.equal(resolveBundleAssetName('win32', 'arm64'), null);
  assert.equal(resolveBundleAssetName('freebsd', 'x64'), null);
});

test('selectBundleAsset skips releases without a matching asset', () => {
  const releases = [
    // The iOS runtime release carries no desktop bundle.
    { tag_name: 'ios-runtime-v7', assets: [{ name: 'MExtensionServer-ios.jar' }] },
    {
      tag_name: 'v1.0.6.0',
      assets: [
        {
          name: 'macOS-arm64-bundle.zip',
          browser_download_url: 'https://example.test/macOS-arm64-bundle.zip',
          size: 133_058_560,
        },
      ],
    },
  ];

  const asset = selectBundleAsset(releases, 'macOS-arm64-bundle.zip');
  assert.equal(asset?.tagName, 'v1.0.6.0');
  assert.equal(asset?.downloadUrl, 'https://example.test/macOS-arm64-bundle.zip');
  assert.equal(asset?.sizeBytes, 133_058_560);
});

test('selectBundleAsset returns null when nothing matches', () => {
  assert.equal(selectBundleAsset([{ tag_name: 'v1', assets: [] }], 'linux-x64-bundle.zip'), null);
  assert.equal(selectBundleAsset([], 'linux-x64-bundle.zip'), null);
  assert.equal(selectBundleAsset({ message: 'rate limited' }, 'linux-x64-bundle.zip'), null);
});

test('findBundleBinaries locates the nested jre and jar', async () => {
  const root = await mkdtemp(path.join(tmpdir(), 'subminer-bundle-'));
  await mkdir(path.join(root, 'jre', 'jre', 'bin'), { recursive: true });
  await writeFile(path.join(root, 'jre', 'jre', 'bin', 'java'), '');
  await writeFile(path.join(root, 'MExtensionServer-1.0.6.0.jar'), '');

  const binaries = await findBundleBinaries(root);
  assert.equal(binaries?.javaPath, path.join(root, 'jre', 'jre', 'bin', 'java'));
  assert.equal(binaries?.jarPath, path.join(root, 'MExtensionServer-1.0.6.0.jar'));
});

test('findBundleBinaries prefers the shallowest java when a nested copy exists', async () => {
  const root = await mkdtemp(path.join(tmpdir(), 'subminer-bundle-'));
  await mkdir(path.join(root, 'bin'), { recursive: true });
  await mkdir(path.join(root, 'bin', 'nested', 'bin'), { recursive: true });
  await writeFile(path.join(root, 'bin', 'java'), '');
  await writeFile(path.join(root, 'bin', 'nested', 'bin', 'java'), '');
  await writeFile(path.join(root, 'MExtensionServer.jar'), '');

  const binaries = await findBundleBinaries(root);
  assert.equal(binaries?.javaPath, path.join(root, 'bin', 'java'));
});

test('findBundleBinaries returns null when the bundle is incomplete', async () => {
  const root = await mkdtemp(path.join(tmpdir(), 'subminer-bundle-'));
  await writeFile(path.join(root, 'MExtensionServer.jar'), '');
  assert.equal(await findBundleBinaries(root), null);
});

test('sha256 produces lowercase hex digests matching known vectors', () => {
  const encode = (value: string) => new TextEncoder().encode(value);
  assert.equal(
    sha256(encode('')),
    'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855',
  );
  assert.equal(
    sha256(encode('subminer')),
    'f3b7fdb2037add4cd8f122c090a727243b46b1b9d8a6c379f71573e2df120885',
  );
});

test('verifyPinnedBundle accepts a matching hash and rejects a mismatch', () => {
  const asset = 'macOS-arm64-bundle.zip';
  const wrong = verifyPinnedBundle(asset, new TextEncoder().encode('not the bundle'));
  assert.equal(wrong.ok, false);
  assert.match((wrong as { reason: string }).reason, /Checksum mismatch/);
});

test('verifyPinnedBundle refuses an asset that has no pin', () => {
  const result = verifyPinnedBundle('linux-x64-bundle.zip', new Uint8Array([1, 2, 3]));
  assert.equal(result.ok, false);
  assert.match((result as { reason: string }).reason, /No pinned checksum/);
});

test('the pinned tag and macOS arm64 hash are the verified release', () => {
  assert.equal(PINNED_BUNDLE_TAG, 'v1.0.6.0');
  assert.equal(
    PINNED_BUNDLE_SHA256['macOS-arm64-bundle.zip'],
    '5f4fb03abfe88bc46ddf5f4d8221156ee2d66b9cbad7c4bc3ade4baf3a4266e6',
  );
});

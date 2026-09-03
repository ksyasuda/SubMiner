import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import {
  BUNDLE_MARKER_FILE,
  bundleReleaseUrl,
  bundleVersionFromJar,
  compareBundleVersions,
  findBundleBinaries,
  MIN_BUNDLE_VERSION,
  parseBundleVersion,
  readBundleMarker,
  resolveBundleAssetName,
  selectBundleAsset,
  systemBundleDirs,
  writeBundleMarker,
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

function release(tag: string, assetNames: string[], extra: Record<string, unknown> = {}) {
  return {
    tag_name: tag,
    assets: assetNames.map((name) => ({
      name,
      browser_download_url: `https://example.test/${tag}/${name}`,
      size: 1,
    })),
    ...extra,
  };
}

test('selectBundleAsset takes the newest release that ships the asset', () => {
  const releases = [
    // The iOS runtime release carries no desktop bundle.
    release('ios-runtime-v7', ['MExtensionServer-ios.jar']),
    release('v1.0.6.1', ['linux-x64-bundle.zip', 'macOS-arm64-bundle.zip']),
    // Older than the one below it: list order must not decide.
    release('v1.0.6.2', ['linux-x64-bundle.zip']),
  ];

  const asset = selectBundleAsset(releases, 'linux-x64-bundle.zip');
  assert.equal(asset?.tagName, 'v1.0.6.2');
  assert.equal(asset?.downloadUrl, 'https://example.test/v1.0.6.2/linux-x64-bundle.zip');

  // The newest release lacks the macOS asset, so the one before it wins there.
  assert.equal(selectBundleAsset(releases, 'macOS-arm64-bundle.zip')?.tagName, 'v1.0.6.1');
});

test('selectBundleAsset skips releases below the minimum, drafts, and prereleases', () => {
  const releases = [
    release('v1.0.5.9', ['linux-x64-bundle.zip']),
    release('v2.0.0.0', ['linux-x64-bundle.zip'], { draft: true }),
    release('v2.0.0.1', ['linux-x64-bundle.zip'], { prerelease: true }),
  ];
  assert.equal(selectBundleAsset(releases, 'linux-x64-bundle.zip'), null);

  releases.push(release(MIN_BUNDLE_VERSION, ['linux-x64-bundle.zip']));
  assert.equal(selectBundleAsset(releases, 'linux-x64-bundle.zip')?.tagName, MIN_BUNDLE_VERSION);
});

test('selectBundleAsset returns null when nothing matches', () => {
  assert.equal(selectBundleAsset([release('v1.0.6.0', [])], 'linux-x64-bundle.zip'), null);
  assert.equal(selectBundleAsset([], 'linux-x64-bundle.zip'), null);
  assert.equal(selectBundleAsset({ message: 'rate limited' }, 'linux-x64-bundle.zip'), null);
});

test('bundleReleaseUrl lists upstream releases', () => {
  assert.equal(
    bundleReleaseUrl(),
    'https://api.github.com/repos/1Selxo/M-Extension-Server/releases?per_page=20',
  );
});

test('parseBundleVersion and compareBundleVersions order tags numerically', () => {
  assert.deepEqual(parseBundleVersion('v1.0.6.2'), [1, 0, 6, 2]);
  assert.deepEqual(parseBundleVersion('1.0.10'), [1, 0, 10]);
  assert.deepEqual(parseBundleVersion('v1.0.6.0-r1'), [1, 0, 6, 0]);
  assert.equal(parseBundleVersion('ios-runtime-v7'), null);
  assert.equal(compareBundleVersions('v1.0.6.2', 'v1.0.6.1'), 1);
  assert.equal(compareBundleVersions('v1.0.10', 'v1.0.9'), 1);
  assert.equal(compareBundleVersions('v1.0.6', 'v1.0.6.0'), 0);
  assert.equal(compareBundleVersions('v1.0.5.9', 'v1.0.6.0'), -1);
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

test('bundleVersionFromJar reads the tag out of the jar name, dropping rebuild suffixes', () => {
  assert.equal(bundleVersionFromJar('/x/MExtensionServer-v1.0.6.0-r1.jar'), 'v1.0.6.0');
  assert.equal(bundleVersionFromJar('/x/MExtensionServer-v1.0.6.2.jar'), 'v1.0.6.2');
  assert.equal(bundleVersionFromJar('/x/MExtensionServer-1.0.6.0.jar'), 'v1.0.6.0');
  assert.equal(bundleVersionFromJar('/x/MExtensionServer.jar'), null);
});

test('bundle marker round-trips and tolerates a missing or malformed file', async () => {
  const root = await mkdtemp(path.join(tmpdir(), 'subminer-bundle-'));
  assert.equal(await readBundleMarker(root), null);
  await writeBundleMarker(root, 'v1.0.6.0');
  assert.equal(await readBundleMarker(root), 'v1.0.6.0');
  await writeFile(path.join(root, BUNDLE_MARKER_FILE), '{not json');
  assert.equal(await readBundleMarker(root), null);
});

test('systemBundleDirs names the AUR package location on Linux only', () => {
  assert.deepEqual(systemBundleDirs('linux'), ['/usr/share/mangatan/extension_server']);
  assert.deepEqual(systemBundleDirs('darwin'), []);
  assert.deepEqual(systemBundleDirs('win32'), []);
});

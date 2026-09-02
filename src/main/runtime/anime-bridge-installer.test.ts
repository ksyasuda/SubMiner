import test from 'node:test';
import assert from 'node:assert/strict';
import { access, mkdir, mkdtemp, readdir, rename, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import {
  BUNDLE_MARKER_FILE,
  readBundleMarker,
  writeBundleMarker,
} from '../../anime-bridge/sidecar-bundle';
import {
  ensureBridgeBinaries,
  findBridgeUpdate,
  stageBridgeUpdate,
  type EnsureBridgeOptions,
} from './anime-bridge-installer';

/** Lay down the upstream bundle shape: a nested jre plus a versioned server jar. */
async function writeBundle(dir: string, jarName: string): Promise<void> {
  await mkdir(path.join(dir, 'jre', 'bin'), { recursive: true });
  await writeFile(path.join(dir, 'jre', 'bin', 'java'), '');
  await writeFile(path.join(dir, jarName), '');
}

async function exists(file: string): Promise<boolean> {
  try {
    await access(file);
    return true;
  } catch {
    return false;
  }
}

const LATEST = 'v1.0.6.2';
const downloadUrl = (tag: string) => `https://example.test/${tag}/linux-x64-bundle.zip`;

/**
 * Upstream's release list answered from memory: a newer release, an older one,
 * an iOS runtime release with no desktop bundle, and one below the minimum.
 */
function fakeUpstream() {
  const calls: string[] = [];
  const bytes = new TextEncoder().encode('zip bytes');
  const asset = (tag: string) => ({
    name: 'linux-x64-bundle.zip',
    browser_download_url: downloadUrl(tag),
    size: bytes.length,
  });
  const fetchImpl: typeof fetch = async (input) => {
    const url = String(input);
    calls.push(url);
    if (url.includes('/releases?')) {
      return new Response(
        JSON.stringify([
          { tag_name: 'ios-runtime-v7', assets: [{ name: 'MExtensionServer-ios.jar' }] },
          { tag_name: 'v1.0.6.1', assets: [asset('v1.0.6.1')] },
          { tag_name: LATEST, assets: [asset(LATEST)] },
          { tag_name: 'v1.0.5.0', assets: [asset('v1.0.5.0')] },
        ]),
      );
    }
    if (url === downloadUrl(LATEST)) return new Response(bytes);
    return new Response('not found', { status: 404 });
  };
  const options = {
    platform: 'linux',
    arch: 'x64',
    fetchImpl,
    extractImpl: async (_zipPath: string, targetDir: string) =>
      writeBundle(targetDir, `MExtensionServer-${LATEST}.jar`),
  } satisfies Partial<EnsureBridgeOptions>;
  return { calls, options };
}

async function tempRoot(): Promise<string> {
  return mkdtemp(path.join(tmpdir(), 'subminer-bridge-install-'));
}

test('anime.bridgeDir wins over every other location and is never downloaded', async () => {
  const root = await tempRoot();
  const configured = path.join(root, 'configured');
  await writeBundle(configured, 'MExtensionServer-v1.0.6.2.jar');
  const system = path.join(root, 'system');
  await writeBundle(system, 'MExtensionServer-v1.0.6.1.jar');
  const { calls, options } = fakeUpstream();

  const install = await ensureBridgeBinaries({
    ...options,
    installDir: path.join(root, 'managed'),
    configuredDir: configured,
    systemDirs: [system],
  });

  assert.equal(install.dir, configured);
  assert.equal(install.origin, 'system');
  assert.equal(install.version, 'v1.0.6.2');
  assert.equal(install.updateAvailable, null);
  assert.equal(install.jarPath, path.join(configured, 'MExtensionServer-v1.0.6.2.jar'));
  assert.deepEqual(calls, []);
});

test('an anime.bridgeDir without a bundle is an error, not a fallback', async () => {
  const root = await tempRoot();
  const { options } = fakeUpstream();
  await assert.rejects(
    ensureBridgeBinaries({
      ...options,
      installDir: path.join(root, 'managed'),
      configuredDir: path.join(root, 'empty'),
      systemDirs: [],
    }),
    /anime\.bridgeDir .*holds no java runtime/,
  );
});

test('a package-manager install is used ahead of the managed copy', async () => {
  const root = await tempRoot();
  const system = path.join(root, 'usr', 'share', 'mangatan', 'extension_server');
  await writeBundle(system, 'MExtensionServer-v1.0.6.2.jar');
  const managed = path.join(root, 'managed');
  await writeBundle(managed, 'MExtensionServer-v1.0.6.0.jar');
  const { options } = fakeUpstream();

  const install = await ensureBridgeBinaries({
    ...options,
    installDir: managed,
    systemDirs: [path.join(root, 'missing'), system],
  });

  assert.equal(install.dir, system);
  assert.equal(install.origin, 'system');
  assert.equal(install.version, 'v1.0.6.2');
});

test('a managed install reports the marker tag over the jar name and is not re-downloaded', async () => {
  const root = await tempRoot();
  const managed = path.join(root, 'managed');
  // The jar name says one thing, the marker another: the marker is what we installed.
  await writeBundle(managed, 'MExtensionServer-v1.0.6.0.jar');
  await writeBundleMarker(managed, 'v1.0.5.0');
  const { calls, options } = fakeUpstream();

  const install = await ensureBridgeBinaries({ ...options, installDir: managed, systemDirs: [] });

  assert.equal(install.origin, 'managed');
  assert.equal(install.version, 'v1.0.5.0');
  assert.deepEqual(calls, []);
});

test('a managed install from before the marker reads its version off the jar', async () => {
  const root = await tempRoot();
  const managed = path.join(root, 'managed');
  await writeBundle(managed, 'MExtensionServer-v1.0.6.0-r1.jar');
  const { options } = fakeUpstream();

  const install = await ensureBridgeBinaries({ ...options, installDir: managed, systemDirs: [] });

  assert.equal(install.version, 'v1.0.6.0');
});

test('with nothing installed the newest release is downloaded and marked', async () => {
  const root = await tempRoot();
  const managed = path.join(root, 'managed');
  const { calls, options } = fakeUpstream();
  const stages: string[] = [];

  const install = await ensureBridgeBinaries({
    ...options,
    installDir: managed,
    systemDirs: [],
    onProgress: (progress) => stages.push(progress.stage),
  });

  assert.equal(install.origin, 'managed');
  assert.equal(install.dir, managed);
  assert.equal(install.version, LATEST);
  assert.equal(calls.length, 2);
  assert.equal(calls[1], downloadUrl(LATEST));
  assert.equal(await readBundleMarker(managed), LATEST);
  assert.deepEqual([...new Set(stages)], ['locating', 'downloading', 'extracting']);
  // The archive is not left behind next to the unpacked bundle.
  assert.ok(!(await readdir(managed)).some((entry) => entry.endsWith('.zip')));
});

test('findBridgeUpdate offers the newest release only to a managed install that is behind it', async () => {
  const { calls, options } = fakeUpstream();

  assert.equal(await findBridgeUpdate({ origin: 'managed', version: 'v1.0.6.0' }, options), LATEST);
  assert.equal(await findBridgeUpdate({ origin: 'managed', version: LATEST }, options), null);
  assert.equal(await findBridgeUpdate({ origin: 'managed', version: 'v1.0.7.0' }, options), null);
  // An unreadable version is offered the newest release as the way back to a known state.
  assert.equal(await findBridgeUpdate({ origin: 'managed', version: null }, options), LATEST);
  assert.equal(calls.length, 4);

  // A system install is pacman's, so upstream is not even asked.
  assert.equal(await findBridgeUpdate({ origin: 'system', version: 'v1.0.0.0' }, options), null);
  assert.equal(calls.length, 4);
});

test('findBridgeUpdate propagates a failed release listing', async () => {
  const fetchImpl: typeof fetch = async () => new Response('rate limited', { status: 403 });
  await assert.rejects(
    findBridgeUpdate(
      { origin: 'managed', version: 'v1.0.6.0' },
      { platform: 'linux', arch: 'x64', fetchImpl },
    ),
    /Could not list anime bridge releases \(403\)/,
  );
});

test('stageBridgeUpdate downloads beside the install and commit swaps it in', async () => {
  const root = await tempRoot();
  const managed = path.join(root, 'managed');
  await writeBundle(managed, 'MExtensionServer-v1.0.5.0.jar');
  await writeBundleMarker(managed, 'v1.0.5.0');
  const { options } = fakeUpstream();

  const staged = await stageBridgeUpdate({ ...options, installDir: managed, systemDirs: [] });

  assert.equal(staged.version, LATEST);
  // The running install is untouched until commit.
  assert.ok(await exists(path.join(managed, 'MExtensionServer-v1.0.5.0.jar')));
  assert.equal(await readBundleMarker(managed), 'v1.0.5.0');
  assert.ok(await exists(path.join(`${managed}.next`, `MExtensionServer-${LATEST}.jar`)));

  const install = await staged.commit();

  assert.equal(install.dir, managed);
  assert.equal(install.version, LATEST);
  assert.equal(install.jarPath, path.join(managed, `MExtensionServer-${LATEST}.jar`));
  assert.ok(!(await exists(path.join(managed, 'MExtensionServer-v1.0.5.0.jar'))));
  assert.ok(!(await exists(`${managed}.next`)));
  assert.ok(await exists(path.join(managed, BUNDLE_MARKER_FILE)));
});

test('stageBridgeUpdate commit leaves the existing install in place when its backup move fails', async () => {
  const root = await tempRoot();
  const managed = path.join(root, 'managed');
  await writeBundle(managed, 'MExtensionServer-v1.0.5.0.jar');
  await writeBundleMarker(managed, 'v1.0.5.0');
  const { options } = fakeUpstream();
  const staged = await stageBridgeUpdate({
    ...options,
    installDir: managed,
    systemDirs: [],
    renameImpl: async () => {
      throw new Error('backup move failed');
    },
  });

  await assert.rejects(staged.commit(), /backup move failed/);

  assert.ok(await exists(path.join(managed, 'MExtensionServer-v1.0.5.0.jar')));
  assert.equal(await readBundleMarker(managed), 'v1.0.5.0');
  assert.ok(await exists(path.join(`${managed}.next`, `MExtensionServer-${LATEST}.jar`)));
  assert.deepEqual(
    (await readdir(root)).filter((entry) => entry.startsWith('managed.backup-')),
    [],
  );
});

test('stageBridgeUpdate commit restores the existing install when activation fails', async () => {
  const root = await tempRoot();
  const managed = path.join(root, 'managed');
  await writeBundle(managed, 'MExtensionServer-v1.0.5.0.jar');
  await writeBundleMarker(managed, 'v1.0.5.0');
  const { options } = fakeUpstream();
  let renameCalls = 0;
  const staged = await stageBridgeUpdate({
    ...options,
    installDir: managed,
    systemDirs: [],
    renameImpl: async (fromPath, toPath) => {
      renameCalls += 1;
      if (renameCalls === 2) throw new Error('activation failed');
      await rename(fromPath, toPath);
    },
  });

  await assert.rejects(staged.commit(), /activation failed/);

  assert.equal(renameCalls, 3);
  assert.ok(await exists(path.join(managed, 'MExtensionServer-v1.0.5.0.jar')));
  assert.equal(await readBundleMarker(managed), 'v1.0.5.0');
  assert.ok(await exists(path.join(`${managed}.next`, `MExtensionServer-${LATEST}.jar`)));
  assert.deepEqual(
    (await readdir(root)).filter((entry) => entry.startsWith('managed.backup-')),
    [],
  );
});

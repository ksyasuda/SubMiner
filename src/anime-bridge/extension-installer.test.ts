import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdir, mkdtemp, readFile, rename, rm, writeFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { installExtension, looksLikeApk, removeExtension } from './extension-installer';
import type { RepoExtension } from './extension-repo';

const PKG = 'eu.kanade.tachiyomi.animeextension.all.example';

function apkBytes(payload = 'APK-BODY'): Uint8Array {
  // APKs are zip archives, so they start with the PK local-file-header magic.
  return new Uint8Array([0x50, 0x4b, 0x03, 0x04, ...new TextEncoder().encode(payload)]);
}

function repoExtension(overrides: Partial<RepoExtension> = {}): RepoExtension {
  return {
    pkg: PKG,
    name: 'Example Source',
    lang: 'all',
    version: '1.2.3',
    versionCode: 12,
    nsfw: false,
    apkUrl: 'https://repo.example/anime/apk/example.apk',
    iconUrl: 'https://repo.example/anime/icon/example.png',
    repoUrl: 'https://repo.example/anime/index.min.json',
    sourceNames: ['Example'],
    ...overrides,
  };
}

function respondWith(bytes: Uint8Array, headers: Record<string, string> = {}): typeof fetch {
  // Uint8Array is a valid Response body at runtime; the DOM lib types disagree.
  const body = bytes as unknown as BodyInit;
  return (async () => new Response(body, { status: 200, headers })) as typeof fetch;
}

test('looksLikeApk accepts the zip magic and rejects anything else', () => {
  assert.equal(looksLikeApk(apkBytes()), true);
  assert.equal(looksLikeApk(new TextEncoder().encode('<!DOCTYPE html>')), false);
  assert.equal(looksLikeApk(new Uint8Array([])), false);
});

test('installExtension writes the apk named after its package', async () => {
  const dir = await mkdtemp(path.join(tmpdir(), 'subminer-install-'));
  const target = await installExtension({
    extensionsDir: dir,
    extension: repoExtension(),
    fetchImpl: respondWith(apkBytes()),
  });

  assert.equal(target, path.join(dir, `${PKG}.apk`));
  assert.match((await readFile(target)).toString(), /APK-BODY/);
});

test('installing again replaces the previous version in place', async () => {
  const dir = await mkdtemp(path.join(tmpdir(), 'subminer-install-'));
  await installExtension({
    extensionsDir: dir,
    extension: repoExtension(),
    fetchImpl: respondWith(apkBytes('OLD')),
  });
  await installExtension({
    extensionsDir: dir,
    extension: repoExtension({ version: '2.0.0', versionCode: 20 }),
    fetchImpl: respondWith(apkBytes('NEW')),
  });

  const contents = (await readFile(path.join(dir, `${PKG}.apk`))).toString();
  assert.match(contents, /NEW/);
  assert.doesNotMatch(contents, /OLD/);
});

test('the extensions directory is created when missing', async () => {
  const root = await mkdtemp(path.join(tmpdir(), 'subminer-install-'));
  const nested = path.join(root, 'does', 'not', 'exist');
  await installExtension({
    extensionsDir: nested,
    extension: repoExtension(),
    fetchImpl: respondWith(apkBytes()),
  });
  assert.equal(existsSync(path.join(nested, `${PKG}.apk`)), true);
});

test('a non-ok response is reported with the extension name', async () => {
  const dir = await mkdtemp(path.join(tmpdir(), 'subminer-install-'));
  const fetchImpl = (async () => new Response('', { status: 404 })) as typeof fetch;

  await assert.rejects(
    () => installExtension({ extensionsDir: dir, extension: repoExtension(), fetchImpl }),
    /Example Source.*404/,
  );
});

test('a response that is not an apk is rejected rather than written', async () => {
  const dir = await mkdtemp(path.join(tmpdir(), 'subminer-install-'));
  // A misconfigured repo commonly serves an HTML error page instead.
  const fetchImpl = respondWith(new TextEncoder().encode('<!DOCTYPE html><html>404</html>'));

  await assert.rejects(
    () => installExtension({ extensionsDir: dir, extension: repoExtension(), fetchImpl }),
    /did not download as an APK/,
  );
  assert.equal(existsSync(path.join(dir, `${PKG}.apk`)), false);
});

test('an oversized download is refused by the declared length', async () => {
  const dir = await mkdtemp(path.join(tmpdir(), 'subminer-install-'));
  const fetchImpl = respondWith(apkBytes(), { 'content-length': '999999999' });

  await assert.rejects(
    () =>
      installExtension({
        extensionsDir: dir,
        extension: repoExtension(),
        fetchImpl,
        maxBytes: 1024,
      }),
    /larger than the 1024 byte limit/,
  );
});

test('an oversized download is refused even when the length header lies', async () => {
  const dir = await mkdtemp(path.join(tmpdir(), 'subminer-install-'));
  const fetchImpl = respondWith(apkBytes('x'.repeat(4096)), { 'content-length': '10' });

  await assert.rejects(
    () =>
      installExtension({
        extensionsDir: dir,
        extension: repoExtension(),
        fetchImpl,
        maxBytes: 1024,
      }),
    /larger than the 1024 byte limit/,
  );
  assert.equal(existsSync(path.join(dir, `${PKG}.apk`)), false);
});

test('the byte limit stops the read instead of buffering the whole body', async () => {
  const dir = await mkdtemp(path.join(tmpdir(), 'subminer-install-'));
  let pushed = 0;
  // Endless body: if the limit were only checked after buffering, this hangs.
  const body = new ReadableStream<Uint8Array>({
    pull(controller) {
      pushed += 1;
      controller.enqueue(new Uint8Array(512));
    },
  });
  const fetchImpl = (async () => new Response(body, { status: 200 })) as typeof fetch;

  await assert.rejects(
    () =>
      installExtension({
        extensionsDir: dir,
        extension: repoExtension(),
        fetchImpl,
        maxBytes: 1024,
      }),
    /larger than the 1024 byte limit/,
  );
  // Only enough chunks to cross the limit were ever read.
  assert.ok(pushed <= 4, `read ${pushed} chunks before aborting`);
});

test('a failed reader cancellation does not hide the size-limit error', async () => {
  const dir = await mkdtemp(path.join(tmpdir(), 'subminer-install-'));
  const body = new ReadableStream<Uint8Array>({
    pull(controller) {
      controller.enqueue(new Uint8Array(1025));
    },
    async cancel() {
      throw new Error('cancel failed');
    },
  });
  const fetchImpl = (async () => new Response(body, { status: 200 })) as typeof fetch;

  await assert.rejects(
    () =>
      installExtension({
        extensionsDir: dir,
        extension: repoExtension(),
        fetchImpl,
        maxBytes: 1024,
      }),
    /larger than the 1024 byte limit/,
  );
});

test('a failed staged write preserves the installed apk and removes the partial file', async () => {
  const dir = await mkdtemp(path.join(tmpdir(), 'subminer-install-'));
  const target = path.join(dir, `${PKG}.apk`);
  await writeFile(target, apkBytes('OLD'));
  let stagedPath = '';

  await assert.rejects(
    () =>
      installExtension({
        extensionsDir: dir,
        extension: repoExtension({ version: '2.0.0', versionCode: 20 }),
        fetchImpl: respondWith(apkBytes('NEW')),
        fileIo: {
          mkdir: (dirPath) => mkdir(dirPath, { recursive: true }),
          async writeFile(filePath, bytes) {
            stagedPath = filePath;
            await writeFile(filePath, bytes.subarray(0, 5));
            throw new Error('simulated disk write failure');
          },
          rename,
          removeFile: (filePath) => rm(filePath, { force: true }),
        },
      }),
    /simulated disk write failure/,
  );

  assert.match((await readFile(target)).toString(), /OLD/);
  assert.notEqual(stagedPath, target);
  assert.equal(existsSync(stagedPath), false);
});

test('a package name carrying path separators cannot escape the extensions dir', async () => {
  const root = await mkdtemp(path.join(tmpdir(), 'subminer-install-'));
  const dir = path.join(root, 'extensions');
  const escaping = `eu.kanade.tachiyomi.animeextension${path.sep}..${path.sep}..${path.sep}pwned`;

  await assert.rejects(
    () =>
      installExtension({
        extensionsDir: dir,
        extension: repoExtension({ pkg: escaping }),
        fetchImpl: respondWith(apkBytes()),
      }),
    /not a valid file name/,
  );
  assert.equal(existsSync(path.join(root, 'pwned.apk')), false);
});

test('removeExtension deletes the file and tolerates a missing one', async () => {
  const dir = await mkdtemp(path.join(tmpdir(), 'subminer-install-'));
  const file = path.join(dir, `${PKG}.apk`);
  await writeFile(file, 'x');

  await removeExtension(dir, PKG);
  assert.equal(existsSync(file), false);
  await removeExtension(dir, PKG);
});

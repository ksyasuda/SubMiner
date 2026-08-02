import test from 'node:test';
import assert from 'node:assert/strict';
import { chmod, mkdtemp, readFile, readdir, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { PreferenceStore } from './preference-store';

async function storeFile(): Promise<string> {
  const dir = await mkdtemp(path.join(tmpdir(), 'subminer-prefs-'));
  return path.join(dir, 'anime-preferences.json');
}

test('values round-trip through the file', async () => {
  const file = await storeFile();
  await new PreferenceStore(file).set('pkg', 'src-1', [{ key: 'address' }]);

  assert.deepEqual(await new PreferenceStore(file).get('pkg', 'src-1'), [{ key: 'address' }]);
});

test('the same bridge source id is isolated between extension packages', async () => {
  const file = await storeFile();
  const store = new PreferenceStore(file);

  await store.set('pkg.one', 'shared-source', [{ key: 'password', value: 'one-secret' }]);
  await store.set('pkg.two', 'shared-source', [{ key: 'password', value: 'two-secret' }]);

  const reloaded = new PreferenceStore(file);
  assert.deepEqual(await reloaded.get('pkg.one', 'shared-source'), [
    { key: 'password', value: 'one-secret' },
  ]);
  assert.deepEqual(await reloaded.get('pkg.two', 'shared-source'), [
    { key: 'password', value: 'two-secret' },
  ]);
});

test('a legacy bare source id is discarded rather than assigned to an unproven package', async () => {
  const file = await storeFile();
  await writeFile(file, JSON.stringify({ 'legacy-source': [{ key: 'address', value: 'saved' }] }));

  const store = new PreferenceStore(file);
  assert.deepEqual(await store.get('pkg.one', 'legacy-source'), []);

  const persisted = JSON.parse(await readFile(file, 'utf8')) as Record<string, unknown>;
  assert.equal(persisted['legacy-source'], undefined);
  assert.equal(persisted['pkg.one:legacy-source'], undefined);
});

test('an ambiguous legacy source id is discarded instead of exposed to either package', async () => {
  const file = await storeFile();
  await writeFile(file, JSON.stringify({ shared: [{ key: 'password', value: 'old-secret' }] }));

  const store = new PreferenceStore(file);
  assert.deepEqual(await store.get('pkg.one', 'shared'), []);
  assert.deepEqual(await store.get('pkg.two', 'shared'), []);

  const persisted = JSON.parse(await readFile(file, 'utf8')) as Record<string, unknown>;
  assert.equal(persisted.shared, undefined);
});

test('concurrent writes on a cold cache do not lose an update', async () => {
  const file = await storeFile();
  const store = new PreferenceStore(file);

  // Both start before either has loaded; unserialized they would each get their
  // own object and the later persist would drop the other's entry.
  await Promise.all([
    store.set('pkg', 'src-1', [{ key: 'a' }]),
    store.set('pkg', 'src-2', [{ key: 'b' }]),
  ]);

  const reloaded = new PreferenceStore(file);
  assert.deepEqual(await reloaded.get('pkg', 'src-1'), [{ key: 'a' }]);
  assert.deepEqual(await reloaded.get('pkg', 'src-2'), [{ key: 'b' }]);
});

test('a clear racing a set is applied in order', async () => {
  const file = await storeFile();
  const store = new PreferenceStore(file);
  await store.set('pkg', 'src', [{ key: 'password' }]);

  await Promise.all([store.clear('pkg'), store.set('other', 'src', [{ key: 'x' }])]);

  const reloaded = new PreferenceStore(file);
  assert.deepEqual(await reloaded.get('pkg', 'src'), []);
  assert.deepEqual(await reloaded.get('other', 'src'), [{ key: 'x' }]);
});

test('the file is written owner-only even when an existing temporary file is permissive', async () => {
  const file = await storeFile();
  await writeFile(`${file}.tmp`, 'stale');
  await chmod(`${file}.tmp`, 0o666);
  await new PreferenceStore(file).set('pkg', 'src-1', [{ key: 'password' }]);

  assert.equal((await stat(file)).mode & 0o777, 0o600);
  assert.deepEqual(await readdir(path.dirname(file)), [path.basename(file)]);
});

test('a corrupt file starts empty rather than blocking the browser', async () => {
  const file = await storeFile();
  await writeFile(file, '{ not json');

  assert.deepEqual(await new PreferenceStore(file).get('pkg', 'src-1'), []);
});

test('malformed persisted values are filtered to preference objects with string keys', async () => {
  const file = await storeFile();
  await writeFile(
    file,
    JSON.stringify({
      'pkg:source': [null, { key: 42 }, 'bad', { key: 'valid', value: 'kept' }],
      'pkg:not-an-array': { key: 'invalid-container' },
    }),
  );

  const store = new PreferenceStore(file);
  assert.deepEqual(await store.get('pkg', 'source'), [{ key: 'valid', value: 'kept' }]);
  assert.deepEqual(await store.get('pkg', 'not-an-array'), []);
});

test('a write replaces the previous contents wholesale', async () => {
  const file = await storeFile();
  const store = new PreferenceStore(file);
  await store.set('pkg', 'src-1', [{ key: 'first' }]);
  await store.set('pkg', 'src-1', [{ key: 'second' }]);

  const parsed = JSON.parse(await readFile(file, 'utf8')) as Record<string, unknown[]>;
  assert.deepEqual(parsed['pkg:src-1'], [{ key: 'second' }]);
});

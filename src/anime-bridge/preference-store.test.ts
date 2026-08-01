import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, readdir, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { PreferenceStore } from './preference-store';

async function storeFile(): Promise<string> {
  const dir = await mkdtemp(path.join(tmpdir(), 'subminer-prefs-'));
  return path.join(dir, 'anime-preferences.json');
}

test('values round-trip through the file', async () => {
  const file = await storeFile();
  await new PreferenceStore(file).set('src-1', [{ key: 'address' }]);

  assert.deepEqual(await new PreferenceStore(file).get('src-1'), [{ key: 'address' }]);
});

test('concurrent writes on a cold cache do not lose an update', async () => {
  const file = await storeFile();
  const store = new PreferenceStore(file);

  // Both start before either has loaded; unserialized they would each get their
  // own object and the later persist would drop the other's entry.
  await Promise.all([store.set('src-1', [{ key: 'a' }]), store.set('src-2', [{ key: 'b' }])]);

  const reloaded = new PreferenceStore(file);
  assert.deepEqual(await reloaded.get('src-1'), [{ key: 'a' }]);
  assert.deepEqual(await reloaded.get('src-2'), [{ key: 'b' }]);
});

test('a clear racing a set is applied in order', async () => {
  const file = await storeFile();
  const store = new PreferenceStore(file);
  await store.set('pkg:src', [{ key: 'password' }]);

  await Promise.all([store.clear('pkg'), store.set('other:src', [{ key: 'x' }])]);

  const reloaded = new PreferenceStore(file);
  assert.deepEqual(await reloaded.get('pkg:src'), []);
  assert.deepEqual(await reloaded.get('other:src'), [{ key: 'x' }]);
});

test('the file is written owner-only and leaves no temporary behind', async () => {
  const file = await storeFile();
  await new PreferenceStore(file).set('src-1', [{ key: 'password' }]);

  const { stat } = await import('node:fs/promises');
  assert.equal((await stat(file)).mode & 0o777, 0o600);
  assert.deepEqual(await readdir(path.dirname(file)), [path.basename(file)]);
});

test('a corrupt file starts empty rather than blocking the browser', async () => {
  const file = await storeFile();
  await writeFile(file, '{ not json');

  assert.deepEqual(await new PreferenceStore(file).get('src-1'), []);
});

test('a write replaces the previous contents wholesale', async () => {
  const file = await storeFile();
  const store = new PreferenceStore(file);
  await store.set('src-1', [{ key: 'first' }]);
  await store.set('src-1', [{ key: 'second' }]);

  const parsed = JSON.parse(await readFile(file, 'utf8')) as Record<string, unknown[]>;
  assert.deepEqual(parsed['src-1'], [{ key: 'second' }]);
});

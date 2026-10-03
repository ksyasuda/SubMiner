import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { resolveHachidoriExtensionPath } from './hachidori-extension';
import { ensureExtensionCopyAsync } from './yomitan-extension-copy';

test('Hachidori resolves development and packaged artifacts without falling back to Yomitan', () => {
  const options = { moduleDir: '/app/dist/core/services', resourcesPath: '/resources' };
  assert.equal(
    resolveHachidoriExtensionPath({
      ...options,
      exists: (p) => p === '/app/build/hachidori/manifest.json',
    }),
    '/app/build/hachidori',
  );
  assert.equal(
    resolveHachidoriExtensionPath({
      ...options,
      exists: (p) => p === '/resources/hachidori/manifest.json',
    }),
    '/resources/hachidori',
  );
  assert.throws(
    () => resolveHachidoriExtensionPath({ ...options, exists: () => false }),
    /build:hachidori/,
  );
});

test('Hachidori updates its own extension copy and preserves the Yomitan copy', async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-backend-copy-'));
  try {
    const source = path.join(root, 'source');
    const profile = path.join(root, 'profile');
    fs.mkdirSync(source);
    fs.writeFileSync(path.join(source, 'manifest.json'), '{"name":"Hachidori","version":"1"}');
    fs.mkdirSync(path.join(profile, 'extensions/yomitan'), { recursive: true });
    fs.writeFileSync(path.join(profile, 'extensions/yomitan/marker'), 'preserved');
    const options = { extensionName: 'hachidori', platform: 'linux' } satisfies Parameters<
      typeof ensureExtensionCopyAsync
    >[2];
    const copy = await ensureExtensionCopyAsync(source, profile, options);
    assert.equal(copy.targetDir, path.join(profile, 'extensions/hachidori'));
    assert.equal((await ensureExtensionCopyAsync(source, profile, options)).copied, false);
    fs.writeFileSync(path.join(source, 'bridge.js'), 'updated');
    assert.equal((await ensureExtensionCopyAsync(source, profile, options)).copied, true);
    assert.equal(
      fs.readFileSync(path.join(profile, 'extensions/yomitan/marker'), 'utf8'),
      'preserved',
    );
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});

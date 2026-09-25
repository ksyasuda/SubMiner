import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { BUNDLED_INTEGRATION_KEYS_FILENAME, readBundledTmdbApiKey } from './bundled-api-key.js';

test('readBundledTmdbApiKey reads the staged key and tolerates a missing or malformed file', () => {
  const distDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-bundled-key-'));
  const filePath = path.join(distDir, BUNDLED_INTEGRATION_KEYS_FILENAME);
  try {
    assert.equal(readBundledTmdbApiKey(distDir), null);
    fs.writeFileSync(filePath, '{"tmdbApiKey":" abc "}');
    assert.equal(readBundledTmdbApiKey(distDir), 'abc');
    fs.writeFileSync(filePath, '{"tmdbApiKey":""}');
    assert.equal(readBundledTmdbApiKey(distDir), null);
    fs.writeFileSync(filePath, 'not json');
    assert.equal(readBundledTmdbApiKey(distDir), null);
  } finally {
    fs.rmSync(distDir, { recursive: true, force: true });
  }
});

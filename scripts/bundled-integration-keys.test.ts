import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  BUNDLED_INTEGRATION_KEYS_FILENAME,
  stageBundledIntegrationKeys,
} from './bundled-integration-keys.mjs';

function withDistDir(work: (distDir: string) => void): void {
  const distDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-bundled-keys-'));
  try {
    work(distDir);
  } finally {
    fs.rmSync(distDir, { recursive: true, force: true });
  }
}

test('stageBundledIntegrationKeys writes the TMDB key from the environment', () => {
  withDistDir((distDir) => {
    const staged = stageBundledIntegrationKeys(distDir, { SUBMINER_TMDB_API_KEY: ' abc123 ' });
    assert.deepEqual(staged, ['tmdb']);
    const written = JSON.parse(
      fs.readFileSync(path.join(distDir, BUNDLED_INTEGRATION_KEYS_FILENAME), 'utf8'),
    );
    assert.deepEqual(written, { tmdbApiKey: 'abc123' });
  });
});

test('stageBundledIntegrationKeys removes a stale file when the variable is unset', () => {
  withDistDir((distDir) => {
    const outputPath = path.join(distDir, BUNDLED_INTEGRATION_KEYS_FILENAME);
    fs.writeFileSync(outputPath, '{"tmdbApiKey":"old"}');
    assert.deepEqual(stageBundledIntegrationKeys(distDir, {}), []);
    assert.equal(fs.existsSync(outputPath), false);
    assert.deepEqual(stageBundledIntegrationKeys(distDir, { SUBMINER_TMDB_API_KEY: '  ' }), []);
  });
});

import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

test('compiled runtime smoke fails clearly when build artifacts are missing', () => {
  const emptyRepo = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-compiled-missing-'));
  try {
    const result = spawnSync(
      process.execPath,
      ['scripts/compiled-runtime-smoke.mjs', '--repo-root', emptyRepo],
      {
        cwd: path.resolve(import.meta.dir, '..'),
        encoding: 'utf8',
      },
    );

    assert.equal(result.status, 1);
    assert.match(result.stderr, /Compiled runtime artifacts are missing/);
    assert.match(result.stderr, /dist\/main-entry\.js/);
    assert.match(result.stderr, /dist\/stats-daemon-runner\.js/);
    assert.doesNotMatch(result.stderr, /must run with Electron/);
  } finally {
    fs.rmSync(emptyRepo, { recursive: true, force: true });
  }
});

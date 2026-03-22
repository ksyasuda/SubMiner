import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import test from 'node:test';

function withTempDir<T>(fn: (dir: string) => T): T {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-patch-modernz-test-'));
  try {
    return fn(dir);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

function writeExecutable(filePath: string, contents: string): void {
  fs.writeFileSync(filePath, contents, 'utf8');
  fs.chmodSync(filePath, 0o755);
}

test('patch-modernz rejects a missing --target value', () => {
  withTempDir((root) => {
    const result = spawnSync('bash', ['scripts/patch-modernz.sh', '--target'], {
      cwd: process.cwd(),
      encoding: 'utf8',
      env: {
        ...process.env,
        HOME: path.join(root, 'home'),
      },
    });

    assert.equal(result.status, 1, result.stderr || result.stdout);
    assert.match(result.stderr, /--target requires a non-empty file path/);
    assert.match(result.stderr, /Usage: patch-modernz\.sh/);
  });
});

test('patch-modernz reports patch failures explicitly', () => {
  withTempDir((root) => {
    const binDir = path.join(root, 'bin');
    const target = path.join(root, 'modernz.lua');
    const patchLog = path.join(root, 'patch.log');

    fs.mkdirSync(binDir, { recursive: true });
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.writeFileSync(target, 'original', 'utf8');

    writeExecutable(
      path.join(binDir, 'patch'),
      `#!/usr/bin/env bash
set -euo pipefail
cat > "${patchLog}"
exit 1
`,
    );

    const result = spawnSync(
      'bash',
      ['scripts/patch-modernz.sh', '--target', target],
      {
        cwd: process.cwd(),
        encoding: 'utf8',
        env: {
          ...process.env,
          HOME: path.join(root, 'home'),
          PATH: `${binDir}:${process.env.PATH || ''}`,
        },
      },
    );

    assert.equal(result.status, 1, result.stderr || result.stdout);
    assert.match(result.stderr, /failed to apply patch to/);
    assert.equal(fs.readFileSync(patchLog, 'utf8').includes('modernz.lua'), true);
  });
});

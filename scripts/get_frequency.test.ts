import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { spawnSync } from 'node:child_process';

function createWorkspace(name: string): string {
  const baseDir = path.join(process.cwd(), '.tmp', 'get-frequency-typecheck-test');
  fs.mkdirSync(baseDir, { recursive: true });
  return fs.mkdtempSync(path.join(baseDir, `${name}-`));
}

test('scripts/get_frequency.ts typechecks in isolation', () => {
  const workspace = createWorkspace('isolated-script');
  const tsconfigPath = path.join(workspace, 'tsconfig.json');

  fs.writeFileSync(
    tsconfigPath,
    JSON.stringify(
      {
        extends: '../../../tsconfig.typecheck.json',
        include: ['../../../scripts/get_frequency.ts'],
        exclude: [],
      },
      null,
      2,
    ),
    'utf8',
  );

  try {
    const result = spawnSync('bunx', ['tsc', '--noEmit', '-p', tsconfigPath], {
      cwd: process.cwd(),
      encoding: 'utf8',
    });

    assert.equal(
      result.status,
      0,
      `expected scripts/get_frequency.ts to typecheck\nstdout:\n${result.stdout}\nstderr:\n${result.stderr}`,
    );
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

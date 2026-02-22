import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { runRuntimeCycleCheck } from './check-runtime-cycles';

type CliResult = {
  status: number | null;
  stdout: string;
  stderr: string;
};

const CHECKER_PATH = path.join(process.cwd(), 'scripts', 'check-runtime-cycles.ts');
const FIXTURE_ROOT = path.join('scripts', 'fixtures', 'runtime-cycles');

function runCheckerCli(args: string[]): CliResult {
  const result = spawnSync('bun', ['run', CHECKER_PATH, ...args], {
    cwd: process.cwd(),
    encoding: 'utf8',
  });

  return {
    status: result.status,
    stdout: result.stdout || '',
    stderr: result.stderr || '',
  };
}

test('acyclic fixture passes runtime cycle check', () => {
  const result = runRuntimeCycleCheck(path.join(FIXTURE_ROOT, 'acyclic'), true);
  assert.equal(result.hasCycle, false);
  assert.equal(result.fileCount >= 3, true);
});

test('cyclic fixture fails via CLI and prints readable cycle path', () => {
  const result = runCheckerCli(['--strict', '--root', path.join(FIXTURE_ROOT, 'cyclic')]);

  assert.equal(result.status, 1);
  assert.match(result.stdout, /^\[FAIL\] runtime cycle check \(strict\)/m);
  assert.match(result.stdout, /module-a\.ts/);
  assert.match(result.stdout, /module-b\.ts|nested\/index\.ts/);
});

test('missing root returns actionable error', () => {
  const missingRoot = path.join(FIXTURE_ROOT, 'does-not-exist');
  const result = runCheckerCli(['--strict', '--root', missingRoot]);

  assert.equal(result.status, 1);
  assert.match(result.stdout, /Runtime root not found:/);
  assert.match(result.stdout, /--root <path>/);
});

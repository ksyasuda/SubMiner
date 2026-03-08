import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const releaseWorkflowPath = resolve(__dirname, '../.github/workflows/release.yml');
const releaseWorkflow = readFileSync(releaseWorkflowPath, 'utf8');

test('publish release leaves prerelease unset so gh creates a normal release', () => {
  assert.ok(!releaseWorkflow.includes('--prerelease'));
});

test('release workflow verifies a committed changelog section before publish', () => {
  assert.match(releaseWorkflow, /bun run changelog:check/);
});

test('release workflow generates release notes from committed changelog output', () => {
  assert.match(releaseWorkflow, /bun run changelog:release-notes/);
  assert.ok(!releaseWorkflow.includes('git log --pretty=format:"- %s"'));
});

import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const ciWorkflowPath = resolve(__dirname, '../.github/workflows/ci.yml');
const ciWorkflow = readFileSync(ciWorkflowPath, 'utf8');

test('ci workflow lints changelog fragments', () => {
  assert.match(ciWorkflow, /bun run changelog:lint/);
});

test('ci workflow checks pull requests for required changelog fragments', () => {
  assert.match(ciWorkflow, /bun run changelog:pr-check/);
  assert.match(ciWorkflow, /skip-changelog/);
});

test('ci workflow verifies generated config examples stay in sync', () => {
  assert.match(ciWorkflow, /bun run verify:config-example/);
});

import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const ciWorkflowPath = resolve(__dirname, '../.github/workflows/ci.yml');
const ciWorkflow = readFileSync(ciWorkflowPath, 'utf8');
const docsPagesWorkflowPath = resolve(__dirname, '../.github/workflows/docs-pages.yml');
const docsPagesWorkflow = readFileSync(docsPagesWorkflowPath, 'utf8');
const packageJsonPath = resolve(__dirname, '../package.json');
const packageJson = JSON.parse(readFileSync(packageJsonPath, 'utf8')) as {
  scripts: Record<string, string>;
};

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

test('package scripts expose a sharded maintained source coverage lane with lcov output', () => {
  assert.equal(
    packageJson.scripts['test:coverage:src'],
    'bun run build:yomitan && bun run scripts/run-coverage-lane.ts bun-src-full --coverage-dir coverage/test-src',
  );
});

test('ci workflow runs the maintained source coverage lane and uploads lcov output', () => {
  assert.match(ciWorkflow, /name: Coverage suite \(maintained source lane\)/);
  assert.match(ciWorkflow, /run: bun run test:coverage:src/);
  assert.match(ciWorkflow, /name: Upload coverage artifact/);
  assert.match(ciWorkflow, /path: coverage\/test-src\/lcov\.info/);
});

test('main docs deploy exists, serializes deploys, and uses Cloudflare credentials', () => {
  assert.match(docsPagesWorkflow, /name: Docs Pages/);
  assert.match(docsPagesWorkflow, /branches:\s*\n\s*-\s*main/);
  assert.match(docsPagesWorkflow, /group:\s*docs-pages-production/);
  assert.match(docsPagesWorkflow, /CLOUDFLARE_API_TOKEN/);
  assert.match(docsPagesWorkflow, /CLOUDFLARE_ACCOUNT_ID/);
  assert.match(docsPagesWorkflow, /CLOUDFLARE_PAGES_PROJECT_NAME/);
  assert.match(docsPagesWorkflow, /pages deploy \.tmp\/docs-versioned-site/);
  assert.match(docsPagesWorkflow, /--branch main/);
});

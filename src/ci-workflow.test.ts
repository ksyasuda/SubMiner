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

test('package scripts expose a sharded maintained source coverage lane with lcov output', () => {
  assert.equal(
    packageJson.scripts['test:coverage:src'],
    'bun run build:yomitan && bun run scripts/run-coverage-lane.ts bun-src-full --coverage-dir coverage/test-src',
  );
});

test('ci delegates its gate instead of duplicating quality steps', () => {
  assert.match(
    ciWorkflow,
    /build-test-audit:\s*\n\s*uses: \.\/\.github\/workflows\/quality-gate\.yml/,
  );
  assert.doesNotMatch(ciWorkflow, /oven-sh\/setup-bun/);
  assert.doesNotMatch(ciWorkflow, /bun run test:coverage:src/);
  assert.doesNotMatch(ciWorkflow, /bun run test:env/);
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

test('docs deploy caches stable archive builds between runs', () => {
  assert.match(docsPagesWorkflow, /actions\/cache@v4/);
  assert.match(docsPagesWorkflow, /\.tmp\/docs-versioned-archive-cache/);
  assert.match(docsPagesWorkflow, /docs-versioned-archives-/);
  assert.match(docsPagesWorkflow, /docs-site\/\.vitepress\/\*\*/);
});

test('docs deploy skips invalid release tags without failing the workflow', () => {
  assert.match(docsPagesWorkflow, /id:\s*tag_guard/);
  assert.match(docsPagesWorkflow, /stable_tag=false/);
  assert.doesNotMatch(docsPagesWorkflow, /exit 78/);
  assert.match(docsPagesWorkflow, /if:\s*steps\.tag_guard\.outputs\.stable_tag != 'false'/);
});

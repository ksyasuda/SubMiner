import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import test from 'node:test';

const repoRoot = process.cwd();

function read(relativePath: string): string {
  return readFileSync(join(repoRoot, relativePath), 'utf8');
}

const requiredDocs = [
  'docs/README.md',
  'docs/architecture/README.md',
  'docs/architecture/domains.md',
  'docs/architecture/layering.md',
  'docs/knowledge-base/README.md',
  'docs/knowledge-base/core-beliefs.md',
  'docs/knowledge-base/catalog.md',
  'docs/knowledge-base/quality.md',
  'docs/workflow/README.md',
  'docs/workflow/planning.md',
  'docs/workflow/verification.md',
] as const;

const metadataFields = ['Status:', 'Last verified:', 'Owner:', 'Read when:'] as const;

test('required internal knowledge-base docs exist', () => {
  for (const relativePath of requiredDocs) {
    assert.equal(existsSync(join(repoRoot, relativePath)), true, `${relativePath} should exist`);
  }
});

test('core internal docs include metadata fields', () => {
  for (const relativePath of requiredDocs) {
    const contents = read(relativePath);
    for (const field of metadataFields) {
      assert.match(contents, new RegExp(field.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')));
    }
  }
});

test('AGENTS.md is a compact map to internal docs', () => {
  const agentsContents = read('AGENTS.md');
  const lineCount = agentsContents.trimEnd().split('\n').length;

  assert.ok(lineCount <= 110, `AGENTS.md should stay compact; got ${lineCount} lines`);
  assert.match(agentsContents, /\.\/docs\/README\.md/);
  assert.match(agentsContents, /\.\/docs\/architecture\/README\.md/);
  assert.match(agentsContents, /\.\/docs\/workflow\/README\.md/);
  assert.match(agentsContents, /\.\/docs\/workflow\/verification\.md/);
  assert.match(agentsContents, /\.\/docs\/knowledge-base\/README\.md/);
  assert.match(agentsContents, /\.\/docs\/RELEASING\.md/);
  assert.match(agentsContents, /`docs-site\/` is user-facing/);
  assert.doesNotMatch(agentsContents, /\.\/docs-site\/development\.md/);
  assert.doesNotMatch(agentsContents, /\.\/docs-site\/architecture\.md/);
});

test('docs-site contributor docs point internal readers to docs/', () => {
  const developmentContents = read('docs-site/development.md');
  const architectureContents = read('docs-site/architecture.md');
  const docsReadmeContents = read('docs-site/README.md');

  assert.match(developmentContents, /docs\/README\.md/);
  assert.match(developmentContents, /docs\/architecture\/README\.md/);
  assert.match(architectureContents, /docs\/architecture\/README\.md/);
  assert.match(docsReadmeContents, /docs\/README\.md/);
});

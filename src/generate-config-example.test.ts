import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import {
  resolveConfigExampleOutputPaths,
  writeConfigExampleArtifacts,
} from './generate-config-example';

function createWorkspace(name: string): string {
  const baseDir = path.join(process.cwd(), '.tmp', 'generate-config-example-test');
  fs.mkdirSync(baseDir, { recursive: true });
  return fs.mkdtempSync(path.join(baseDir, `${name}-`));
}

test('resolveConfigExampleOutputPaths includes sibling docs repo and never local docs/public', () => {
  const workspace = createWorkspace('with-docs-repo');
  const projectRoot = path.join(workspace, 'SubMiner');
  const docsRepoRoot = path.join(workspace, 'subminer-docs');

  fs.mkdirSync(projectRoot, { recursive: true });
  fs.mkdirSync(docsRepoRoot, { recursive: true });

  try {
    const outputPaths = resolveConfigExampleOutputPaths({ cwd: projectRoot });

    assert.deepEqual(outputPaths, [
      path.join(projectRoot, 'config.example.jsonc'),
      path.join(docsRepoRoot, 'public', 'config.example.jsonc'),
    ]);
    assert.equal(
      outputPaths.includes(path.join(projectRoot, 'docs', 'public', 'config.example.jsonc')),
      false,
    );
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('resolveConfigExampleOutputPaths stays repo-local when sibling docs repo is absent', () => {
  const workspace = createWorkspace('without-docs-repo');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(projectRoot, { recursive: true });

  try {
    const outputPaths = resolveConfigExampleOutputPaths({ cwd: projectRoot });

    assert.deepEqual(outputPaths, [path.join(projectRoot, 'config.example.jsonc')]);
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('writeConfigExampleArtifacts creates parent directories for resolved outputs', () => {
  const workspace = createWorkspace('write-artifacts');
  const projectRoot = path.join(workspace, 'SubMiner');
  const docsRepoRoot = path.join(workspace, 'subminer-docs');
  const template = '{\n  "ok": true\n}\n';

  fs.mkdirSync(projectRoot, { recursive: true });
  fs.mkdirSync(docsRepoRoot, { recursive: true });

  try {
    const writtenPaths = writeConfigExampleArtifacts(template, {
      cwd: projectRoot,
      deps: { log: () => {} },
    });

    assert.deepEqual(writtenPaths, [
      path.join(projectRoot, 'config.example.jsonc'),
      path.join(docsRepoRoot, 'public', 'config.example.jsonc'),
    ]);
    assert.equal(fs.readFileSync(path.join(projectRoot, 'config.example.jsonc'), 'utf8'), template);
    assert.equal(
      fs.readFileSync(path.join(docsRepoRoot, 'public', 'config.example.jsonc'), 'utf8'),
      template,
    );
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

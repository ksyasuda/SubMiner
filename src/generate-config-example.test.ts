import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import {
  generateConfigExampleTemplate,
  resolveConfigExampleOutputPaths,
  writeConfigExampleArtifacts,
} from './generate-config-example';

function createWorkspace(name: string): string {
  const baseDir = path.join(process.cwd(), '.tmp', 'generate-config-example-test');
  fs.mkdirSync(baseDir, { recursive: true });
  return fs.mkdtempSync(path.join(baseDir, `${name}-`));
}

test('resolveConfigExampleOutputPaths includes in-repo docs site and never local docs/public', () => {
  const workspace = createWorkspace('with-docs-site');
  const projectRoot = path.join(workspace, 'SubMiner');
  const docsSiteRoot = path.join(projectRoot, 'docs-site');

  fs.mkdirSync(projectRoot, { recursive: true });
  fs.mkdirSync(docsSiteRoot, { recursive: true });

  try {
    const outputPaths = resolveConfigExampleOutputPaths({ cwd: projectRoot });

    assert.deepEqual(outputPaths, [
      path.join(projectRoot, 'config.example.jsonc'),
      path.join(docsSiteRoot, 'public', 'config.example.jsonc'),
    ]);
    assert.equal(
      outputPaths.includes(path.join(projectRoot, 'docs', 'public', 'config.example.jsonc')),
      false,
    );
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('resolveConfigExampleOutputPaths stays repo-local when docs site is absent', () => {
  const workspace = createWorkspace('without-docs-site');
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
  const docsSiteRoot = path.join(projectRoot, 'docs-site');
  const template = '{\n  "ok": true\n}\n';

  fs.mkdirSync(projectRoot, { recursive: true });
  fs.mkdirSync(docsSiteRoot, { recursive: true });

  try {
    const writtenPaths = writeConfigExampleArtifacts(template, {
      cwd: projectRoot,
      deps: { log: () => {} },
    });

    assert.deepEqual(writtenPaths, [
      path.join(projectRoot, 'config.example.jsonc'),
      path.join(docsSiteRoot, 'public', 'config.example.jsonc'),
    ]);
    assert.equal(fs.readFileSync(path.join(projectRoot, 'config.example.jsonc'), 'utf8'), template);
    assert.equal(
      fs.readFileSync(path.join(docsSiteRoot, 'public', 'config.example.jsonc'), 'utf8'),
      template,
    );
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('generateConfigExampleTemplate uses the canonical example socket path', () => {
  const template = generateConfigExampleTemplate();

  assert.match(template, /"socketPath": "\/tmp\/subminer-socket"/);
});

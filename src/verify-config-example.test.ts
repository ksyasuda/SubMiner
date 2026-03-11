import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import {
  verifyConfigExampleArtifacts,
  type ConfigExampleVerificationResult,
} from './verify-config-example';

function createWorkspace(name: string): string {
  const baseDir = path.join(process.cwd(), '.tmp', 'verify-config-example-test');
  fs.mkdirSync(baseDir, { recursive: true });
  return fs.mkdtempSync(path.join(baseDir, `${name}-`));
}

function assertResult(
  result: ConfigExampleVerificationResult,
  expected: {
    missingPaths?: string[];
    stalePaths?: string[];
  },
) {
  assert.deepEqual(result.missingPaths, expected.missingPaths ?? []);
  assert.deepEqual(result.stalePaths, expected.stalePaths ?? []);
}

test('verifyConfigExampleArtifacts reports repo config example when missing', () => {
  const workspace = createWorkspace('missing-root');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(projectRoot, { recursive: true });

  try {
    const result = verifyConfigExampleArtifacts({ cwd: projectRoot });

    assertResult(result, {
      missingPaths: [path.join(projectRoot, 'config.example.jsonc')],
    });
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('verifyConfigExampleArtifacts reports stale docs-site artifact when docs site exists', () => {
  const workspace = createWorkspace('stale-docs-site');
  const projectRoot = path.join(workspace, 'SubMiner');
  const docsSiteRoot = path.join(projectRoot, 'docs-site');

  fs.mkdirSync(projectRoot, { recursive: true });
  fs.mkdirSync(path.join(docsSiteRoot, 'public'), { recursive: true });
  fs.writeFileSync(path.join(projectRoot, 'config.example.jsonc'), 'fresh\n');
  fs.writeFileSync(path.join(docsSiteRoot, 'public', 'config.example.jsonc'), 'stale\n');

  try {
    const result = verifyConfigExampleArtifacts({
      cwd: projectRoot,
      template: 'fresh\n',
    });

    assertResult(result, {
      stalePaths: [path.join(docsSiteRoot, 'public', 'config.example.jsonc')],
    });
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('verifyConfigExampleArtifacts passes when repo and docs-site artifacts match', () => {
  const workspace = createWorkspace('matching-artifacts');
  const projectRoot = path.join(workspace, 'SubMiner');
  const docsSiteRoot = path.join(projectRoot, 'docs-site');
  const template = '{\n  "ok": true\n}\n';

  fs.mkdirSync(projectRoot, { recursive: true });
  fs.mkdirSync(path.join(docsSiteRoot, 'public'), { recursive: true });
  fs.writeFileSync(path.join(projectRoot, 'config.example.jsonc'), template);
  fs.writeFileSync(path.join(docsSiteRoot, 'public', 'config.example.jsonc'), template);

  try {
    const result = verifyConfigExampleArtifacts({
      cwd: projectRoot,
      template,
    });

    assertResult(result, {});
    assert.deepEqual(result.outputPaths, [
      path.join(projectRoot, 'config.example.jsonc'),
      path.join(docsSiteRoot, 'public', 'config.example.jsonc'),
    ]);
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

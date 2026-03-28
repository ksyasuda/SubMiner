import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import test from 'node:test';

const repoRoot = process.cwd();
const classifyScript = path.join(
  repoRoot,
  '.agents/skills/subminer-change-verification/scripts/classify_subminer_diff.sh',
);
const verifyScript = path.join(
  repoRoot,
  '.agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh',
);

function withTempDir<T>(fn: (dir: string) => T): T {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-change-verification-test-'));
  try {
    return fn(dir);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

function runBash(args: string[]) {
  return spawnSync('bash', args, {
    cwd: repoRoot,
    env: process.env,
    encoding: 'utf8',
  });
}

function parseArtifactDir(stdout: string): string {
  const match = stdout.match(/^artifacts: (.+)$/m);
  assert.ok(match, `expected artifact_dir in stdout, got:\n${stdout}`);
  return match[1] ?? '';
}

function readSummaryJson(artifactDir: string) {
  return JSON.parse(fs.readFileSync(path.join(artifactDir, 'summary.json'), 'utf8')) as {
    sessionId: string;
    status: string;
    lanes: string[];
    blockers?: string[];
    artifactDir: string;
    pathSelectionMode?: string;
    steps: Array<{
      lane: string;
      name: string;
      stdout: string;
      stderr: string;
      note: string;
    }>;
  };
}

test('classifier marks launcher and plugin paths as real-runtime candidates', () => {
  const result = runBash([classifyScript, 'launcher/mpv.ts', 'plugin/subminer/process.lua']);

  assert.equal(result.status, 0, result.stderr || result.stdout);
  assert.match(result.stdout, /^lane:launcher-plugin$/m);
  assert.match(result.stdout, /^flag:real-runtime-candidate$/m);
  assert.doesNotMatch(result.stdout, /real-gui-candidate/);
});

test('verifier blocks requested real-runtime lane when runtime execution is not allowed', () => {
  withTempDir((root) => {
    const artifactDir = path.join(root, 'artifacts');
    const result = runBash([
      verifyScript,
      '--dry-run',
      '--artifact-dir',
      artifactDir,
      '--lane',
      'real-runtime',
      'launcher/mpv.ts',
    ]);

    assert.equal(result.status, 0, result.stdout);

    const summary = readSummaryJson(artifactDir);
    assert.equal(summary.status, 'blocked');
    assert.deepEqual(summary.lanes, ['real-runtime']);
    assert.ok(summary.sessionId.length > 0);
    assert.ok(summary.blockers?.some((entry) => entry.includes('--allow-real-runtime')));
    assert.equal(fs.existsSync(path.join(artifactDir, 'summary.json')), true);
  });
});

test('verifier fails closed for unknown lanes', () => {
  withTempDir((root) => {
    const artifactDir = path.join(root, 'artifacts');
    const result = runBash([
      verifyScript,
      '--dry-run',
      '--artifact-dir',
      artifactDir,
      '--lane',
      'not-a-lane',
      'src/main.ts',
    ]);

    assert.equal(result.status, 0, result.stdout);

    const summary = readSummaryJson(artifactDir);
    assert.equal(summary.status, 'blocked');
    assert.deepEqual(summary.lanes, ['not-a-lane']);
    assert.ok(summary.blockers?.some((entry) => entry.includes('unknown lane')));
  });
});

test('verifier keeps non-passing step artifacts distinct across lanes', () => {
  withTempDir((root) => {
    const artifactDir = path.join(root, 'artifacts');
    const result = runBash([
      verifyScript,
      '--dry-run',
      '--artifact-dir',
      artifactDir,
      '--lane',
      'docs',
      '--lane',
      'not-a-lane',
      'src/main.ts',
    ]);

    assert.equal(result.status, 0, result.stdout);

    const summary = readSummaryJson(artifactDir);
    const docsStep = summary.steps.find((step) => step.lane === 'docs' && step.name === 'docs-kb');
    const unknownStep = summary.steps.find(
      (step) => step.lane === 'not-a-lane' && step.name === 'unknown-lane',
    );

    assert.ok(docsStep);
    assert.ok(unknownStep);
    assert.notEqual(docsStep?.stdout, unknownStep?.stdout);
    assert.equal(fs.existsSync(path.join(artifactDir, docsStep!.stdout)), true);
    assert.equal(fs.existsSync(path.join(artifactDir, unknownStep!.stdout)), true);
  });
});

test('verifier records the real-runtime lease blocker once', () => {
  withTempDir((root) => {
    const artifactDir = path.join(root, 'artifacts');
    const leaseDir = path.join(
      repoRoot,
      '.tmp',
      'skill-verification',
      'locks',
      'exclusive-real-runtime',
    );
    fs.mkdirSync(leaseDir, { recursive: true });
    fs.writeFileSync(path.join(leaseDir, 'session_id'), 'other-session');

    try {
      const result = runBash([
        verifyScript,
        '--dry-run',
        '--artifact-dir',
        artifactDir,
        '--allow-real-runtime',
        '--lane',
        'real-runtime',
        'launcher/mpv.ts',
      ]);

      assert.equal(result.status, 0, result.stdout);

      const summary = readSummaryJson(artifactDir);
      assert.deepEqual(summary.blockers, ['real-runtime lease already held by other-session']);
    } finally {
      fs.rmSync(leaseDir, { recursive: true, force: true });
    }
  });
});

test('verifier allocates unique session ids and artifact roots by default', () => {
  const first = runBash([verifyScript, '--dry-run', '--lane', 'core', 'src/main.ts']);
  const second = runBash([verifyScript, '--dry-run', '--lane', 'core', 'src/main.ts']);

  assert.equal(first.status, 0, first.stderr || first.stdout);
  assert.equal(second.status, 0, second.stderr || second.stdout);

  const firstArtifactDir = parseArtifactDir(first.stdout);
  const secondArtifactDir = parseArtifactDir(second.stdout);

  try {
    const firstSummary = readSummaryJson(firstArtifactDir);
    const secondSummary = readSummaryJson(secondArtifactDir);

    assert.notEqual(firstSummary.sessionId, secondSummary.sessionId);
    assert.notEqual(firstArtifactDir, secondArtifactDir);
    assert.equal(firstSummary.pathSelectionMode, 'explicit-lanes');
    assert.equal(secondSummary.pathSelectionMode, 'explicit-lanes');
  } finally {
    fs.rmSync(firstArtifactDir, { recursive: true, force: true });
    fs.rmSync(secondArtifactDir, { recursive: true, force: true });
  }
});

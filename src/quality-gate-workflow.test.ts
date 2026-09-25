import test from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const qualityGateWorkflowPath = resolve(__dirname, '../.github/workflows/quality-gate.yml');
const qualityGateWorkflow = existsSync(qualityGateWorkflowPath)
  ? readFileSync(qualityGateWorkflowPath, 'utf8').replace(/\r\n/g, '\n')
  : '';
const verificationDocPath = resolve(__dirname, '../docs/workflow/verification.md');
const verificationDoc = readFileSync(verificationDocPath, 'utf8').replace(/\r\n/g, '\n');

test('quality gate is a reusable workflow', () => {
  assert.match(qualityGateWorkflow, /on:\s*\n\s*workflow_call:/);
  assert.match(qualityGateWorkflow, /permissions:\s*\n\s*contents: read/);
});

test('quality gate checkout does not persist GitHub credentials', () => {
  assert.match(
    qualityGateWorkflow,
    /uses: actions\/checkout@v4[\s\S]*?fetch-depth: 0[\s\S]*?submodules: true[\s\S]*?persist-credentials: false/,
  );
});

test('quality gate runs non-covered source suites and lets coverage gate the src lane', () => {
  assert.match(qualityGateWorkflow, /name: Install Lua/);
  assert.match(
    qualityGateWorkflow,
    /apt_sources=\(-o Dir::Etc::sourcelist=sources\.list\.d\/ubuntu\.sources -o Dir::Etc::sourceparts=-\)/,
  );
  assert.match(qualityGateWorkflow, /apt-get\s+"\$\{apt_sources\[@\]\}"\s+update/);
  assert.match(qualityGateWorkflow, /apt-get\s+"\$\{apt_sources\[@\]\}"\s+install\s+-y\s+lua5\.4/);
  assert.match(
    qualityGateWorkflow,
    /Launcher unit and script suites\n\s*run: bun run test:launcher:unit:src && bun run test:scripts/,
  );
  assert.doesNotMatch(qualityGateWorkflow, /bun run test:fast/);
  assert.match(qualityGateWorkflow, /run: bun run test:coverage:src/);
});

test('quality gate runs launcher smoke once through the environment suite and keeps artifacts', () => {
  assert.match(qualityGateWorkflow, /name: Environment suite\n\s*run: bun run test:env/);
  assert.doesNotMatch(qualityGateWorkflow, /run: bun run test:launcher:smoke:src/);
  assert.match(
    qualityGateWorkflow,
    /name: Upload launcher smoke artifacts \(on failure\)[\s\S]*?if: failure\(\)[\s\S]*?path: \.tmp\/launcher-smoke\/\*\*/,
  );
});

test('quality gate uploads maintained source coverage', () => {
  assert.match(qualityGateWorkflow, /run: bun run test:coverage:src/);
  assert.match(qualityGateWorkflow, /name: Upload coverage artifact/);
  assert.match(qualityGateWorkflow, /path: coverage\/test-src\/lcov\.info/);
});

test('quality gate preserves stats, compiled SQLite, and dist runtime checks', () => {
  assert.match(qualityGateWorkflow, /run: bun run test:stats/);
  assert.match(qualityGateWorkflow, /run: bun run build/);
  assert.match(qualityGateWorkflow, /run: bun run test:immersion:sqlite:dist/);
  assert.match(qualityGateWorkflow, /run: bun run test:smoke:dist/);
});

test('quality gate keeps pull request changelog enforcement event-aware', () => {
  assert.match(qualityGateWorkflow, /bun run changelog:lint/);
  assert.match(qualityGateWorkflow, /if: github\.event_name == 'pull_request'/);
  assert.match(
    qualityGateWorkflow,
    /env:\s*\n\s*BASE_REF: \${{ github\.base_ref }}\s*\n\s*PR_LABELS: \${{ join\(github\.event\.pull_request\.labels\.\*\.name, ','\) }}\s*\n\s*run: bun run changelog:pr-check --base-ref "origin\/\$BASE_REF" --head-ref "HEAD" --labels "\$PR_LABELS"/,
  );
  assert.match(qualityGateWorkflow, /skip-changelog/);
});

test('quality gate verifies generated config examples', () => {
  assert.match(qualityGateWorkflow, /bun run verify:config-example/);
});

test('quality gate blocks on high-severity audit findings', () => {
  const securityAuditStep =
    qualityGateWorkflow.match(/- name: Security audit[\s\S]*?(?=\n {6}- name:|\n?$)/)?.[0] ?? '';
  const dependencyAuditPolicySection =
    verificationDoc.match(/## Dependency Audit Policy[\s\S]*?(?=\n## |\n?$)/)?.[0] ?? '';

  assert.match(securityAuditStep, /run: bun audit --audit-level high/);
  assert.doesNotMatch(securityAuditStep, /continue-on-error: true/);
  assert.match(
    dependencyAuditPolicySection,
    /`bun audit --audit-level high` blocks the reusable quality gate/,
  );
});

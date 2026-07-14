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

test('quality gate installs Lua and runs the environment suite before coverage', () => {
  assert.match(qualityGateWorkflow, /name: Install Lua/);
  assert.match(qualityGateWorkflow, /apt-get install -y lua5\.4/);
  assert.match(
    qualityGateWorkflow,
    /Test suite \(source\)\n\s*run: bun run test:fast\n\s*\n\s*- name: Environment suite\n\s*run: bun run test:env\n\s*\n\s*- name: Coverage suite \(maintained source lane\)/,
  );
});

test('quality gate uploads maintained source coverage', () => {
  assert.match(qualityGateWorkflow, /run: bun run test:coverage:src/);
  assert.match(qualityGateWorkflow, /name: Upload coverage artifact/);
  assert.match(qualityGateWorkflow, /path: coverage\/test-src\/lcov\.info/);
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

test('quality gate documents its temporary non-blocking audit policy', () => {
  assert.match(
    qualityGateWorkflow,
    /name: Security audit\s*\n\s*# Dependency audit remains advisory; see docs\/workflow\/verification\.md\.\s*\n\s*run: bun audit --audit-level high\s*\n\s*continue-on-error: true/,
  );
  assert.match(
    verificationDoc,
    /## Dependency Audit Policy[\s\S]*`bun audit --audit-level high` remains advisory[\s\S]*Remove\s+`continue-on-error` only after the audit passes/,
  );
});

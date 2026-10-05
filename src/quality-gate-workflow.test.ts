import test from 'node:test';
import assert from 'node:assert/strict';
import { resolve } from 'node:path';
import { jobSteps, readWorkflow, stepRunsCommand } from './workflow-test-helpers';

const qualityGateWorkflow = readWorkflow(
  resolve(__dirname, '../.github/workflows/quality-gate.yml'),
);

test('quality gate blocks on high-severity audit findings', () => {
  const auditStep = jobSteps(qualityGateWorkflow, 'quality-gate').find((step) =>
    stepRunsCommand(step, /^bun audit --audit-level high\b/),
  );
  assert.ok(auditStep, 'quality gate must run bun audit --audit-level high');
  assert.equal(auditStep['continue-on-error'] ?? false, false);
});

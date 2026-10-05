import test from 'node:test';
import assert from 'node:assert/strict';
import { resolve } from 'node:path';
import { jobSteps, readWorkflow, stepRunsCommand } from './workflow-test-helpers';

const workflowPath = (file: string): string => resolve(__dirname, '../.github/workflows', file);

test('each release build job installs stats dependencies before packaging', () => {
  const workflow = readWorkflow(workflowPath('package-release.yml'));
  for (const job of ['build-linux', 'build-macos', 'build-windows']) {
    const steps = jobSteps(workflow, job);
    const install = steps.findIndex((step) =>
      step.run?.includes('cd stats && bun install --frozen-lockfile'),
    );
    const build = steps.findIndex((step) =>
      /bun run build:(appimage|mac|win)/.test(step.run ?? ''),
    );
    assert(install >= 0 && build > install, `${job} must install stats before packaging`);
  }
});

test('stable release verifies the committed changelog before creating or editing the release', () => {
  const steps = jobSteps(readWorkflow(workflowPath('release.yml')), 'release');
  const checkIndex = steps.findIndex((step) =>
    stepRunsCommand(step, /^bun run changelog:check --version "\$RELEASE_VERSION"/),
  );
  const publishIndex = steps.findIndex((step) =>
    stepRunsCommand(step, /^gh release (create|edit)\b/),
  );

  assert.notEqual(checkIndex, -1);
  assert.notEqual(publishIndex, -1);
  assert.ok(checkIndex < publishIndex);
});

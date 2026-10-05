import test from 'node:test';
import assert from 'node:assert/strict';
import { readdirSync } from 'node:fs';
import { resolve } from 'node:path';
import {
  readWorkflow,
  stepsReadingUndeclaredEnv,
  templateExpressionsInRunBodies,
} from './workflow-test-helpers';

// Policy lints that apply to every workflow, so new workflows are covered automatically.
const workflowsDir = resolve(__dirname, '../.github/workflows');
const workflows = readdirSync(workflowsDir)
  .filter((file) => /\.ya?ml$/.test(file))
  .sort()
  .map((file) => ({ file, workflow: readWorkflow(resolve(workflowsDir, file)) }));

test('workflow policy lint finds the workflow files', () => {
  assert.ok(workflows.length > 0);
});

for (const { file, workflow } of workflows) {
  test(`${file} passes values to shell through env, not template expressions`, () => {
    assert.deepEqual(templateExpressionsInRunBodies(workflow), []);
    assert.deepEqual(stepsReadingUndeclaredEnv(workflow), []);
  });
}

test('reusable workflows default to read-only contents and do not persist checkout credentials', () => {
  const reusable = workflows.filter(
    ({ workflow }) => workflow.on && 'workflow_call' in workflow.on,
  );
  assert.ok(reusable.length > 0);

  for (const { file, workflow } of reusable) {
    assert.deepEqual(workflow.permissions, { contents: 'read' }, file);
    for (const [job, definition] of Object.entries(workflow.jobs ?? {})) {
      for (const step of definition?.steps ?? []) {
        if (step.uses?.startsWith('actions/checkout@')) {
          assert.equal(step.with?.['persist-credentials'], false, `${file}/${job}`);
        }
      }
    }
  }
});

import test from 'node:test';
import assert from 'node:assert/strict';
import { resolve } from 'node:path';
import { commandPositions, jobSteps, readWorkflow, stepRunsCommand } from './workflow-test-helpers';

const workflowPath = (file: string): string => resolve(__dirname, '../.github/workflows', file);
const prereleaseWorkflow = readWorkflow(workflowPath('prerelease.yml'));
const packageWorkflow = readWorkflow(workflowPath('package-release.yml'));

test('release callers pass only the declared packaging secrets', () => {
  const signingSecrets = [
    'CSC_LINK',
    'CSC_KEY_PASSWORD',
    'APPLE_ID',
    'APPLE_APP_SPECIFIC_PASSWORD',
    'APPLE_TEAM_ID',
  ];
  // The bundled TMDB key is optional: artifacts stay valid without it and
  // users fall back to their own tmdb.apiKey.
  const optionalSecrets = ['SUBMINER_TMDB_API_KEY'];
  assert.deepEqual(packageWorkflow.on?.workflow_call?.secrets, {
    ...Object.fromEntries(signingSecrets.map((name) => [name, { required: true }])),
    ...Object.fromEntries(optionalSecrets.map((name) => [name, { required: false }])),
  });
  for (const workflow of [prereleaseWorkflow, readWorkflow(workflowPath('release.yml'))]) {
    assert.equal(workflow.jobs?.package?.uses, './.github/workflows/package-release.yml');
    assert.deepEqual(
      workflow.jobs?.package?.secrets,
      Object.fromEntries(
        [...signingSecrets, ...optionalSecrets].map((name) => [
          name,
          '${{ secrets.' + name + ' }}',
        ]),
      ),
    );
  }
});

test('prerelease workflow rejects committed notes generated for a different beta or rc', () => {
  // Matched at command positions only, so commenting the check out or quoting it
  // inside an echo fails the test rather than silently satisfying it.
  const steps = jobSteps(prereleaseWorkflow, 'release');
  const checkIndex = steps.findIndex((step) =>
    stepRunsCommand(
      step,
      /^bun run changelog:check-prerelease-notes --version "\$RELEASE_VERSION"/,
    ),
  );
  const publishIndex = steps.findIndex((step) =>
    stepRunsCommand(step, /^gh release (create|edit)\b/),
  );

  assert.notEqual(checkIndex, -1);
  assert.notEqual(publishIndex, -1);
  // Stale notes are already published if the check runs after the release.
  assert.ok(checkIndex < publishIndex);
});

test('prerelease stays a draft until every asset is uploaded', () => {
  const commands = jobSteps(prereleaseWorkflow, 'release')
    .flatMap((step) => commandPositions(step))
    .filter((command) => /^gh release (create|edit|upload)\b/.test(command));
  const uploads = commands.flatMap((command, index) =>
    command.startsWith('gh release upload') ? [index] : [],
  );
  const lastUpload = Math.max(-1, ...uploads);
  const undraft = commands.findIndex((command) => /\s--draft=false\b/.test(command));

  assert.notEqual(lastUpload, -1);
  assert.ok(undraft > lastUpload, 'the release must only be undrafted after the last upload');
  for (const command of commands.slice(0, lastUpload)) {
    if (/^gh release (create|edit)\b/.test(command)) {
      assert.match(command, /\s--draft(\s|$)/, `published before upload: ${command}`);
    }
  }
});

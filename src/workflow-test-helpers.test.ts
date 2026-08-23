import test from 'node:test';
import assert from 'node:assert/strict';
import {
  commandPositions,
  executableRunLines,
  stepRunsCommand,
  stepsMissingEnvDeclaration,
  templateExpressionsInRunBodies,
} from './workflow-test-helpers';

const runs = (run: string): boolean => stepRunsCommand({ run }, /^bun run verify --flag "\$VALUE"/);

test('stepRunsCommand matches a command that actually executes', () => {
  assert.equal(runs('bun run verify --flag "$VALUE"'), true);
  assert.equal(runs('if ! bun run verify --flag "$VALUE"; then\nexit 1\nfi'), true);
  assert.equal(runs('set -e && bun run verify --flag "$VALUE"'), true);
  assert.equal(runs('  bun run verify --flag "$VALUE" || exit 1'), true);
});

test('stepRunsCommand rejects commands that are only mentioned, not run', () => {
  assert.equal(runs('# bun run verify --flag "$VALUE"'), false);
  assert.equal(runs('echo \'bun run verify --flag "$VALUE"\''), false);
  assert.equal(runs("printf '%s\\n' 'bun run verify --flag \"$VALUE\"'"), false);
  assert.equal(runs('echo "run: bun run verify --flag \\"$VALUE\\"" >> notes.txt'), false);
  // A different argument list is a different command.
  assert.equal(runs('bun run verify'), false);
});

test('stepRunsCommand ignores separators inside quotes and inline comments', () => {
  assert.equal(runs('echo \'note; bun run verify --flag "$VALUE"\''), false);
  assert.equal(runs('echo "note && bun run verify --flag \\"$VALUE\\""'), false);
  assert.equal(runs("printf '%s\\n' 'a | bun run verify --flag \"$VALUE\"'"), false);
  assert.equal(runs('if false; then # bun run verify --flag "$VALUE"'), false);
  // A trailing comment does not hide the command in front of it.
  assert.equal(runs('bun run verify --flag "$VALUE" # keep this'), true);
  // A pipe is a real separator; a redirect is not.
  assert.equal(runs('cat notes | bun run verify --flag "$VALUE"'), true);
  assert.equal(stepRunsCommand({ run: 'gh release view "$V" 2>&1 | tee log' }, /^tee\b/), true);
});

test('stepRunsCommand treats backslash-escaped separators as literal text', () => {
  assert.equal(runs(String.raw`echo foo \; bun run verify --flag "$VALUE"`), false);
  assert.equal(runs(String.raw`echo foo \| bun run verify --flag "$VALUE"`), false);
  assert.equal(runs(String.raw`find . -exec bun run verify --flag "$VALUE" \;`), false);
  // An escape does not swallow a following real separator.
  assert.equal(runs(String.raw`echo a\b; bun run verify --flag "$VALUE"`), true);
});

test('commandPositions splits on separators and strips control-flow prefixes', () => {
  assert.deepEqual(
    commandPositions({ run: 'if gh release view "$V"; then\ngh release edit "$V"\nfi' }),
    ['gh release view "$V"', 'then', 'gh release edit "$V"', 'fi'],
  );
});

test('executableRunLines drops blank and comment-only lines', () => {
  assert.deepEqual(executableRunLines({ run: '\n# a comment\n  \nreal command\n' }), [
    'real command',
  ]);
});

test('templateExpressionsInRunBodies reports every expression spelling in a run body', () => {
  const workflow = {
    jobs: {
      release: {
        steps: [
          { name: 'Safe', env: { V: '${{ steps.version.outputs.VERSION }}' }, run: 'echo "$V"' },
          { name: 'Dotted', run: 'echo "${{ steps.version.outputs.VERSION }}"' },
          { name: 'Bracketed', run: 'echo "${{ steps.version.outputs[\'VERSION\'] }}"' },
          { name: 'Github', run: 'echo "${{ github[\'ref_name\'] }}"' },
        ],
      },
    },
  };

  assert.deepEqual(templateExpressionsInRunBodies(workflow), [
    'release/Dotted: ${{ steps.version.outputs.VERSION }}',
    "release/Bracketed: ${{ steps.version.outputs['VERSION'] }}",
    "release/Github: ${{ github['ref_name'] }}",
  ]);
});

test('stepsMissingEnvDeclaration finds shell reads with no matching env entry', () => {
  const workflow = {
    jobs: {
      release: {
        steps: [
          { name: 'Declared', env: { TAG: 'x' }, run: 'echo "$TAG"' },
          { name: 'Undeclared', run: 'echo "${TAG}"' },
          { name: 'Unrelated', run: 'echo "$TAGGED"' },
        ],
      },
    },
  };

  assert.deepEqual(stepsMissingEnvDeclaration(workflow, 'TAG'), ['release/Undeclared']);
});

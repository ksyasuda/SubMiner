import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const releaseWorkflowPath = resolve(__dirname, '../.github/workflows/release.yml');
const releaseWorkflow = readFileSync(releaseWorkflowPath, 'utf8');
const makefilePath = resolve(__dirname, '../Makefile');
const makefile = readFileSync(makefilePath, 'utf8');
const packageJsonPath = resolve(__dirname, '../package.json');
const packageJson = JSON.parse(readFileSync(packageJsonPath, 'utf8')) as {
  scripts: Record<string, string>;
};

test('publish release leaves prerelease unset so gh creates a normal release', () => {
  assert.ok(!releaseWorkflow.includes('--prerelease'));
});

test('publish release forces an existing draft tag release to become public', () => {
  assert.ok(releaseWorkflow.includes('--draft=false'));
});

test('release workflow verifies a committed changelog section before publish', () => {
  assert.match(releaseWorkflow, /bun run changelog:check/);
});

test('release workflow generates release notes from committed changelog output', () => {
  assert.match(releaseWorkflow, /bun run changelog:release-notes/);
  assert.ok(!releaseWorkflow.includes('git log --pretty=format:"- %s"'));
});

test('release workflow includes the Windows installer in checksums and uploaded assets', () => {
  assert.match(releaseWorkflow, /files=\(release\/\*\.AppImage release\/\*\.dmg release\/\*\.exe release\/\*\.zip release\/\*\.tar\.gz dist\/launcher\/subminer\)/);
  assert.match(releaseWorkflow, /artifacts=\([\s\S]*release\/\*\.exe[\s\S]*release\/SHA256SUMS\.txt[\s\S]*\)/);
});

test('release package scripts disable implicit electron-builder publishing', () => {
  assert.match(packageJson.scripts['build:appimage'] ?? '', /--publish never/);
  assert.match(packageJson.scripts['build:mac'] ?? '', /--publish never/);
  assert.match(packageJson.scripts['build:win'] ?? '', /--publish never/);
});

test('windows release workflow retries SignPath submission and fails only after exhausting attempts', () => {
  assert.match(releaseWorkflow, /Submit Windows signing request \(attempt 1\)/);
  assert.match(releaseWorkflow, /Submit Windows signing request \(attempt 2\)/);
  assert.match(releaseWorkflow, /Submit Windows signing request \(attempt 3\)/);
  assert.match(releaseWorkflow, /All SignPath signing attempts failed; rerun the workflow when SignPath is healthy\./);
});

test('Makefile routes Windows install-plugin setup through bun and documents Windows builds', () => {
  assert.match(makefile, /windows\) printf '%s\\n' "\[INFO\] Windows builds run via: bun run build:win" ;;/);
  assert.match(makefile, /bun \.\/scripts\/configure-plugin-binary-path\.mjs/);
});

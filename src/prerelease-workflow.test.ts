import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import {
  jobSteps,
  readWorkflow,
  stepRunsCommand,
  stepsMissingEnvDeclaration,
  templateExpressionsInRunBodies,
} from './workflow-test-helpers';

const prereleaseWorkflowPath = resolve(__dirname, '../.github/workflows/prerelease.yml');
const prereleaseWorkflow = readFileSync(prereleaseWorkflowPath, 'utf8').replace(/\r\n/g, '\n');
const parsedPrereleaseWorkflow = readWorkflow(prereleaseWorkflowPath);
const packageJsonPath = resolve(__dirname, '../package.json');
const packageJson = JSON.parse(readFileSync(packageJsonPath, 'utf8')) as {
  scripts: Record<string, string>;
};

test('prerelease workflow triggers on beta and rc tags only', () => {
  assert.match(prereleaseWorkflow, /name: Prerelease/);
  const tagsBlock = prereleaseWorkflow.match(/tags:\s*\n((?:\s*-\s*'[^']+'\s*\n?)+)/);
  assert.ok(tagsBlock, 'Workflow tags block not found');
  const tagsText = tagsBlock[1];
  assert.ok(tagsText, 'Workflow tags entries not found');
  const tagPatterns = [...tagsText.matchAll(/-\s*'([^']+)'/g)].map(([, pattern]) => pattern);
  assert.deepEqual(tagPatterns, ['v*-beta.*', 'v*-rc.*']);
});

test('package scripts expose prerelease notes generation separately from stable changelog build', () => {
  assert.equal(
    packageJson.scripts['changelog:prerelease-notes'],
    'bun run scripts/build-changelog.ts prerelease-notes',
  );
});

test('prerelease workflow uses committed prerelease notes and never calls claude in CI', () => {
  assert.match(prereleaseWorkflow, /--notes-file release\/prerelease-notes\.md/);
  assert.doesNotMatch(prereleaseWorkflow, /run: bun run changelog:prerelease-notes/);
  assert.doesNotMatch(prereleaseWorkflow, /run: bun run changelog:build/);
});

test('prerelease delegates its quality gate instead of duplicating quality steps', () => {
  assert.match(
    prereleaseWorkflow,
    /quality-gate:\s*\n\s*permissions:\s*\n\s*contents: read\s*\n\s*uses: \.\/\.github\/workflows\/quality-gate\.yml/,
  );
  const qualityGateJob = prereleaseWorkflow.match(/quality-gate:[\s\S]*?(?=\n  build-linux:)/)?.[0];
  assert.ok(qualityGateJob);
  assert.doesNotMatch(qualityGateJob, /oven-sh\/setup-bun/);
  assert.doesNotMatch(qualityGateJob, /bun run test:coverage:src/);
  assert.doesNotMatch(qualityGateJob, /bun run test:env/);
});

test('prerelease workflow publishes GitHub prereleases and keeps them off latest', () => {
  assert.match(prereleaseWorkflow, /gh release edit[\s\S]*--prerelease/);
  assert.match(prereleaseWorkflow, /gh release create[\s\S]*--prerelease/);
  assert.match(prereleaseWorkflow, /gh release create[\s\S]*--latest=false/);
});

test('prerelease packaging workflows scope dependency caches by runner architecture', () => {
  const archScopedCacheKeyMatches = prereleaseWorkflow.match(
    /key:\s*\${{\s*runner\.os\s*}}-\${{\s*runner\.arch\s*}}-bun-/g,
  );
  const archScopedRestoreKeyMatches = prereleaseWorkflow.match(
    /\${{\s*runner\.os\s*}}-\${{\s*runner\.arch\s*}}-bun-/g,
  );
  assert.equal(archScopedCacheKeyMatches?.length, 4);
  assert.ok((archScopedRestoreKeyMatches?.length ?? 0) >= 8);
});

test('prerelease workflow builds and uploads all release platforms', () => {
  assert.match(prereleaseWorkflow, /build-linux:/);
  assert.match(prereleaseWorkflow, /build-macos:/);
  assert.match(prereleaseWorkflow, /build-windows:/);
  assert.match(prereleaseWorkflow, /name: appimage/);
  assert.match(prereleaseWorkflow, /name: macos/);
  assert.match(prereleaseWorkflow, /name: windows/);
});

test('prerelease workflow publishes the same release assets as the stable workflow', () => {
  assert.match(
    prereleaseWorkflow,
    /files=\(release\/\*\.AppImage release\/\*\.dmg release\/\*\.exe release\/\*\.zip release\/\*\.tar\.gz release\/latest\*\.yml release\/\*\.blockmap dist\/launcher\/subminer\)/,
  );
  assert.match(
    prereleaseWorkflow,
    /artifacts=\([\s\S]*release\/\*\.exe[\s\S]*release\/latest\*\.yml[\s\S]*release\/\*\.blockmap[\s\S]*release\/SHA256SUMS\.txt[\s\S]*\)/,
  );
});

test('prerelease workflow uploads updater metadata without builder debug YAML files', () => {
  assert.match(prereleaseWorkflow, /release\/latest\*\.yml/);
  assert.doesNotMatch(prereleaseWorkflow, /release\/\*\.yml/);
});

test('prerelease workflow writes checksum entries using release asset basenames', () => {
  assert.match(prereleaseWorkflow, /: > release\/SHA256SUMS\.txt/);
  assert.match(prereleaseWorkflow, /for file in "\$\{files\[@\]\}"; do/);
  assert.match(prereleaseWorkflow, /\$\{file##\*\/\}/);
  assert.doesNotMatch(
    prereleaseWorkflow,
    /sha256sum "\$\{files\[@\]\}" > release\/SHA256SUMS\.txt/,
  );
});

test('prerelease workflow relies on builder artifact names without post-build zip renames', () => {
  assert.doesNotMatch(prereleaseWorkflow, /Rename Windows ZIP artifacts/);
  assert.doesNotMatch(prereleaseWorkflow, /Rename-Item[\s\S]*-win\.zip/);
});

test('prerelease workflow validates artifacts before publishing the release and only undrafts after upload', () => {
  const artifactsIndex = prereleaseWorkflow.indexOf('artifacts=(');
  const createIndex = prereleaseWorkflow.indexOf('gh release create');
  const uploadIndex = prereleaseWorkflow.indexOf('gh release upload');
  const undraftIndex = prereleaseWorkflow.indexOf('--draft=false');

  assert.notEqual(artifactsIndex, -1);
  assert.notEqual(createIndex, -1);
  assert.notEqual(uploadIndex, -1);
  assert.notEqual(undraftIndex, -1);
  assert.ok(artifactsIndex < createIndex);
  assert.ok(uploadIndex < undraftIndex);
  assert.match(prereleaseWorkflow, /gh release create[\s\S]*--draft[\s\S]*--prerelease/);
});

test('prerelease workflow does not publish to AUR', () => {
  assert.doesNotMatch(prereleaseWorkflow, /aur-publish:/);
  assert.doesNotMatch(prereleaseWorkflow, /AUR_SSH_PRIVATE_KEY/);
  assert.doesNotMatch(prereleaseWorkflow, /scripts\/update-aur-package\.sh/);
});

test('prerelease workflow rejects committed notes generated for a different beta or rc', () => {
  assert.equal(
    packageJson.scripts['changelog:check-prerelease-notes'],
    'bun run scripts/build-changelog.ts check-prerelease-notes',
  );

  // Matched against executable lines only, so commenting the check out fails the
  // test rather than silently satisfying it.
  const steps = jobSteps(parsedPrereleaseWorkflow, 'release');
  const checkIndex = steps.findIndex((step) =>
    stepRunsCommand(step, /changelog:check-prerelease-notes --version "\$RELEASE_VERSION"/),
  );
  const publishIndex = steps.findIndex((step) => stepRunsCommand(step, /gh release (create|edit)/));

  assert.notEqual(checkIndex, -1);
  assert.notEqual(publishIndex, -1);
  // Stale notes are already published if the check runs after the release.
  assert.ok(checkIndex < publishIndex);
});

test('prerelease workflow keeps tag-derived values out of shell bodies', () => {
  assert.deepEqual(templateExpressionsInRunBodies(parsedPrereleaseWorkflow), []);
  assert.deepEqual(stepsMissingEnvDeclaration(parsedPrereleaseWorkflow, 'RELEASE_VERSION'), []);
});

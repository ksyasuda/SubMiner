import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const prereleaseWorkflowPath = resolve(__dirname, '../.github/workflows/prerelease.yml');
const prereleaseWorkflow = readFileSync(prereleaseWorkflowPath, 'utf8').replace(/\r\n/g, '\n');
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

test('prerelease workflow generates prerelease notes from pending fragments', () => {
  assert.match(prereleaseWorkflow, /bun run changelog:prerelease-notes --version/);
  assert.doesNotMatch(prereleaseWorkflow, /bun run changelog:build --version/);
});

test('prerelease workflow includes the environment suite in the gate sequence', () => {
  assert.match(
    prereleaseWorkflow,
    /Test suite \(source\)\n\s*run: bun run test:fast\n\s*\n\s*- name: Environment suite(?: \(source\))?\n\s*run: bun run test:env\n\s*\n\s*- name: Coverage suite \(maintained source lane\)/,
  );
});

test('prerelease workflow publishes GitHub prereleases and keeps them off latest', () => {
  assert.match(prereleaseWorkflow, /gh release edit[\s\S]*--prerelease/);
  assert.match(prereleaseWorkflow, /gh release create[\s\S]*--prerelease/);
  assert.match(prereleaseWorkflow, /gh release create[\s\S]*--latest=false/);
});

test('prerelease workflow scopes dependency caches by runner architecture', () => {
  const archScopedCacheKeyMatches = prereleaseWorkflow.match(
    /key:\s*\${{\s*runner\.os\s*}}-\${{\s*runner\.arch\s*}}-bun-/g,
  );
  const archScopedRestoreKeyMatches = prereleaseWorkflow.match(
    /\${{\s*runner\.os\s*}}-\${{\s*runner\.arch\s*}}-bun-/g,
  );
  assert.equal(archScopedCacheKeyMatches?.length, 5);
  assert.ok((archScopedRestoreKeyMatches?.length ?? 0) >= 10);
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
    /files=\(release\/\*\.AppImage release\/\*\.dmg release\/\*\.exe release\/\*\.zip release\/\*\.tar\.gz dist\/launcher\/subminer\)/,
  );
  assert.match(
    prereleaseWorkflow,
    /artifacts=\([\s\S]*release\/\*\.exe[\s\S]*release\/SHA256SUMS\.txt[\s\S]*\)/,
  );
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

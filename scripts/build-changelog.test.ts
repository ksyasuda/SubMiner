import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';

async function loadModule() {
  return import('./build-changelog');
}

function createWorkspace(name: string): string {
  const baseDir = path.join(process.cwd(), '.tmp', 'build-changelog-test');
  fs.mkdirSync(baseDir, { recursive: true });
  return fs.mkdtempSync(path.join(baseDir, `${name}-`));
}

test('resolveChangelogOutputPaths stays repo-local and never writes docs paths', async () => {
  const { resolveChangelogOutputPaths } = await loadModule();
  const workspace = createWorkspace('with-docs-repo');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(projectRoot, { recursive: true });

  try {
    const outputPaths = resolveChangelogOutputPaths({ cwd: projectRoot });

    assert.deepEqual(outputPaths, [path.join(projectRoot, 'CHANGELOG.md')]);
    assert.equal(outputPaths.includes(path.join(projectRoot, 'docs', 'changelog.md')), false);
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('writeChangelogArtifacts ignores README, groups fragments by type, writes release notes, and deletes only fragment files', async () => {
  const { writeChangelogArtifacts } = await loadModule();
  const workspace = createWorkspace('write-artifacts');
  const projectRoot = path.join(workspace, 'SubMiner');
  const existingChangelog = [
    '# Changelog',
    '',
    '## v0.4.0 (2026-03-01)',
    '- Existing fix',
    '',
  ].join('\n');

  fs.mkdirSync(projectRoot, { recursive: true });
  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(path.join(projectRoot, 'CHANGELOG.md'), existingChangelog, 'utf8');
  fs.writeFileSync(
    path.join(projectRoot, 'changes', 'README.md'),
    '# Changelog Fragments\n\nIgnored helper text.\n',
    'utf8',
  );
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: added', 'area: overlay', '', '- Added release fragments.'].join('\n'),
    'utf8',
  );
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '002.md'),
    ['type: fixed', 'area: release', '', 'Fixed release notes generation.'].join('\n'),
    'utf8',
  );

  try {
    const result = writeChangelogArtifacts({
      cwd: projectRoot,
      version: '0.4.1',
      date: '2026-03-07',
    });

    assert.deepEqual(result.outputPaths, [path.join(projectRoot, 'CHANGELOG.md')]);
    assert.deepEqual(result.deletedFragmentPaths, [
      path.join(projectRoot, 'changes', '001.md'),
      path.join(projectRoot, 'changes', '002.md'),
    ]);
    assert.equal(fs.existsSync(path.join(projectRoot, 'changes', '001.md')), false);
    assert.equal(fs.existsSync(path.join(projectRoot, 'changes', '002.md')), false);
    assert.equal(fs.existsSync(path.join(projectRoot, 'changes', 'README.md')), true);

    const changelog = fs.readFileSync(path.join(projectRoot, 'CHANGELOG.md'), 'utf8');
    assert.match(
      changelog,
      /^# Changelog\n\n## v0\.4\.1 \(2026-03-07\)\n\n### Added\n- Overlay: Added release fragments\.\n\n### Fixed\n- Release: Fixed release notes generation\.\n\n## v0\.4\.0 \(2026-03-01\)\n- Existing fix\n$/m,
    );

    const releaseNotes = fs.readFileSync(
      path.join(projectRoot, 'release', 'release-notes.md'),
      'utf8',
    );
    assert.match(releaseNotes, /## Highlights\n### Added\n- Overlay: Added release fragments\./);
    assert.match(releaseNotes, /### Fixed\n- Release: Fixed release notes generation\./);
    assert.match(releaseNotes, /## Installation\n\nSee the README and docs\/installation guide/);
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('verifyChangelogReadyForRelease ignores README but rejects pending fragments and missing version sections', async () => {
  const { verifyChangelogReadyForRelease } = await loadModule();
  const workspace = createWorkspace('verify-release');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(path.join(projectRoot, 'CHANGELOG.md'), '# Changelog\n', 'utf8');
  fs.writeFileSync(
    path.join(projectRoot, 'changes', 'README.md'),
    '# Changelog Fragments\n',
    'utf8',
  );
  fs.writeFileSync(path.join(projectRoot, 'changes', '001.md'), '- Pending fragment.\n', 'utf8');

  try {
    assert.throws(
      () => verifyChangelogReadyForRelease({ cwd: projectRoot, version: '0.4.1' }),
      /Pending changelog fragments/,
    );

    fs.rmSync(path.join(projectRoot, 'changes', '001.md'));

    assert.throws(
      () => verifyChangelogReadyForRelease({ cwd: projectRoot, version: '0.4.1' }),
      /Missing CHANGELOG section for v0\.4\.1/,
    );
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('verifyChangelogReadyForRelease rejects explicit release versions that do not match package.json', async () => {
  const { verifyChangelogReadyForRelease } = await loadModule();
  const workspace = createWorkspace('verify-release-version-match');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(
    path.join(projectRoot, 'package.json'),
    JSON.stringify({ name: 'subminer', version: '0.4.0' }, null, 2),
    'utf8',
  );
  fs.writeFileSync(
    path.join(projectRoot, 'CHANGELOG.md'),
    '# Changelog\n\n## v0.4.1 (2026-03-09)\n- Ready.\n',
    'utf8',
  );

  try {
    assert.throws(
      () => verifyChangelogReadyForRelease({ cwd: projectRoot, version: '0.4.1' }),
      /package\.json version \(0\.4\.0\) does not match requested release version \(0\.4\.1\)/,
    );
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('verifyChangelogFragments rejects invalid metadata', async () => {
  const { verifyChangelogFragments } = await loadModule();
  const workspace = createWorkspace('lint-invalid');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: nope', 'area: overlay', '', '- Invalid type.'].join('\n'),
    'utf8',
  );

  try {
    assert.throws(
      () => verifyChangelogFragments({ cwd: projectRoot }),
      /must declare type as one of/,
    );
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('verifyPullRequestChangelog requires fragments for user-facing changes and skips docs-only changes', async () => {
  const { verifyPullRequestChangelog } = await loadModule();

  assert.throws(
    () =>
      verifyPullRequestChangelog({
        changedEntries: [{ path: 'src/main-entry.ts', status: 'M' }],
        changedLabels: [],
      }),
    /requires a changelog fragment/,
  );

  assert.doesNotThrow(() =>
    verifyPullRequestChangelog({
      changedEntries: [{ path: 'docs/RELEASING.md', status: 'M' }],
      changedLabels: [],
    }),
  );

  assert.doesNotThrow(() =>
    verifyPullRequestChangelog({
      changedEntries: [{ path: 'src/main-entry.ts', status: 'M' }],
      changedLabels: ['skip-changelog'],
    }),
  );

  assert.throws(
    () =>
      verifyPullRequestChangelog({
        changedEntries: [
          { path: 'src/main-entry.ts', status: 'M' },
          { path: 'changes/001.md', status: 'D' },
        ],
        changedLabels: [],
      }),
    /requires a changelog fragment/,
  );

  assert.doesNotThrow(() =>
    verifyPullRequestChangelog({
      changedEntries: [
        { path: 'src/main-entry.ts', status: 'M' },
        { path: 'changes/001.md', status: 'A' },
      ],
      changedLabels: [],
    }),
  );
});

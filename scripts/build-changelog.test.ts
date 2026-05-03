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

type RunClaudeArgs = { input: string; args: string[] };

function recordingRunClaude(responder: (input: string) => string): {
  runClaude: (input: string, args: string[]) => string;
  calls: RunClaudeArgs[];
} {
  const calls: RunClaudeArgs[] = [];
  return {
    calls,
    runClaude(input, args) {
      calls.push({ input, args });
      return responder(input);
    },
  };
}

function modeFromPrompt(input: string): 'changelog' | 'release-notes' | null {
  // Anchor to start-of-line so we don't accidentally match the instructions text,
  // which mentions "MODE: changelog" and "MODE: release-notes" mid-sentence.
  const match = /^MODE: (changelog|release-notes)$/m.exec(input);
  return (match?.[1] as 'changelog' | 'release-notes') ?? null;
}

function fragmentTypesInPrompt(input: string): string[] {
  return input
    .split(/\r?\n/)
    .filter((line) => line.startsWith('type: '))
    .map((line) => line.slice('type: '.length).trim());
}

function defaultPolishedBody(input: string): string {
  const mode = modeFromPrompt(input);
  const types = fragmentTypesInPrompt(input);
  const sections: string[] = [];

  const has = (t: string) => types.includes(t);
  const hasBreaking = /^breaking: true$/m.test(input);
  if (hasBreaking) {
    sections.push('### Breaking Changes\n- Polished: breaking change.');
  }
  if (has('added')) {
    sections.push('### Added\n- Polished: added entry.');
  }
  if (has('changed')) {
    sections.push('### Changed\n- Polished: changed entry.');
  }
  if (has('fixed')) {
    sections.push('### Fixed\n- Polished: fixed entry.');
  }
  if (has('docs')) {
    sections.push('### Docs\n- Polished: docs entry.');
  }
  if (mode === 'changelog' && has('internal')) {
    sections.push(
      '<details>\n<summary>Internal changes</summary>\n\n### Internal\n- Polished: internal entry.\n\n</details>',
    );
  }

  if (sections.length === 0) {
    sections.push('### Changed\n- Polished: empty fallback.');
  }

  return sections.join('\n\n');
}

function defaultStubClaude() {
  return recordingRunClaude(defaultPolishedBody);
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
    const stub = defaultStubClaude();
    const result = writeChangelogArtifacts({
      cwd: projectRoot,
      version: '0.4.1',
      date: '2026-03-07',
      deps: { runClaude: stub.runClaude },
    });

    assert.deepEqual(result.outputPaths, [path.join(projectRoot, 'CHANGELOG.md')]);
    assert.deepEqual(result.deletedFragmentPaths, [
      path.join(projectRoot, 'changes', '001.md'),
      path.join(projectRoot, 'changes', '002.md'),
    ]);
    assert.equal(fs.existsSync(path.join(projectRoot, 'changes', '001.md')), false);
    assert.equal(fs.existsSync(path.join(projectRoot, 'changes', '002.md')), false);
    assert.equal(fs.existsSync(path.join(projectRoot, 'changes', 'README.md')), true);

    assert.equal(
      stub.calls.length,
      2,
      'expected one Claude call per output (changelog + release notes)',
    );
    assert.deepEqual(
      stub.calls.map((call) => modeFromPrompt(call.input)),
      ['changelog', 'release-notes'],
    );

    const changelog = fs.readFileSync(path.join(projectRoot, 'CHANGELOG.md'), 'utf8');
    assert.match(changelog, /^# Changelog\n\n## v0\.4\.1 \(2026-03-07\)\n\n/);
    assert.match(changelog, /### Added\n- Polished: added entry\./);
    assert.match(changelog, /### Fixed\n- Polished: fixed entry\./);
    assert.match(changelog, /## v0\.4\.0 \(2026-03-01\)\n- Existing fix\n$/);

    const releaseNotes = fs.readFileSync(
      path.join(projectRoot, 'release', 'release-notes.md'),
      'utf8',
    );
    assert.match(releaseNotes, /## Highlights\n### Added\n- Polished: added entry\./);
    assert.match(releaseNotes, /### Fixed\n- Polished: fixed entry\./);
    assert.match(releaseNotes, /## Installation\n\nSee the README and docs\/installation guide/);
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('writeChangelogArtifacts skips changelog prepend when release section already exists', async () => {
  const { writeChangelogArtifacts } = await loadModule();
  const workspace = createWorkspace('write-artifacts-existing-version');
  const projectRoot = path.join(workspace, 'SubMiner');
  const existingChangelog = [
    '# Changelog',
    '',
    '## v0.4.1 (2026-03-07)',
    '### Added',
    '- Existing release bullet.',
    '',
  ].join('\n');

  fs.mkdirSync(projectRoot, { recursive: true });
  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(path.join(projectRoot, 'CHANGELOG.md'), existingChangelog, 'utf8');
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: added', 'area: overlay', '', '- Stale release fragment.'].join('\n'),
    'utf8',
  );

  try {
    const result = writeChangelogArtifacts({
      cwd: projectRoot,
      version: '0.4.1',
      date: '2026-03-08',
    });

    assert.deepEqual(result.deletedFragmentPaths, [path.join(projectRoot, 'changes', '001.md')]);
    assert.equal(fs.existsSync(path.join(projectRoot, 'changes', '001.md')), false);

    const changelog = fs.readFileSync(path.join(projectRoot, 'CHANGELOG.md'), 'utf8');
    assert.equal(changelog, existingChangelog);
    const releaseNotes = fs.readFileSync(
      path.join(projectRoot, 'release', 'release-notes.md'),
      'utf8',
    );
    assert.match(releaseNotes, /## Highlights\n### Added\n- Existing release bullet\./);
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('writeStableReleaseArtifacts reuses the requested version and date for changelog, release notes, and docs-site output', async () => {
  const { writeStableReleaseArtifacts } = await loadModule();
  const workspace = createWorkspace('write-stable-release-artifacts');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.mkdirSync(path.join(projectRoot, 'docs-site'), { recursive: true });
  fs.writeFileSync(
    path.join(projectRoot, 'package.json'),
    JSON.stringify({ name: 'subminer', version: '0.4.1' }, null, 2),
    'utf8',
  );
  fs.writeFileSync(path.join(projectRoot, 'CHANGELOG.md'), '# Changelog\n', 'utf8');
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: fixed', 'area: release', '', '- Reused explicit stable release date.'].join('\n'),
    'utf8',
  );

  try {
    const stub = defaultStubClaude();
    const result = writeStableReleaseArtifacts({
      cwd: projectRoot,
      version: '0.4.1',
      date: '2026-03-07',
      deps: { runClaude: stub.runClaude },
    });

    assert.deepEqual(result.outputPaths, [path.join(projectRoot, 'CHANGELOG.md')]);
    assert.equal(result.releaseNotesPath, path.join(projectRoot, 'release', 'release-notes.md'));
    assert.equal(result.docsChangelogPath, path.join(projectRoot, 'docs-site', 'changelog.md'));

    const changelog = fs.readFileSync(path.join(projectRoot, 'CHANGELOG.md'), 'utf8');
    const docsChangelog = fs.readFileSync(
      path.join(projectRoot, 'docs-site', 'changelog.md'),
      'utf8',
    );

    assert.match(changelog, /## v0\.4\.1 \(2026-03-07\)/);
    assert.match(docsChangelog, /## v0\.4\.1 \(2026-03-07\)/);
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

test('writeChangelogArtifacts renders breaking changes section above type sections', async () => {
  const { writeChangelogArtifacts } = await loadModule();
  const workspace = createWorkspace('breaking-changes');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(projectRoot, { recursive: true });
  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(path.join(projectRoot, 'CHANGELOG.md'), '# Changelog\n', 'utf8');
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: changed', 'area: config', 'breaking: true', '', '- Renamed `foo` to `bar`.'].join('\n'),
    'utf8',
  );
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '002.md'),
    ['type: fixed', 'area: overlay', '', '- Fixed subtitle rendering.'].join('\n'),
    'utf8',
  );

  try {
    const stub = defaultStubClaude();
    writeChangelogArtifacts({
      cwd: projectRoot,
      version: '0.5.0',
      date: '2026-04-06',
      deps: { runClaude: stub.runClaude },
    });

    const changelog = fs.readFileSync(path.join(projectRoot, 'CHANGELOG.md'), 'utf8');
    const breakingIndex = changelog.indexOf('### Breaking Changes');
    const changedIndex = changelog.indexOf('### Changed');
    const fixedIndex = changelog.indexOf('### Fixed');

    assert.notEqual(breakingIndex, -1, 'Breaking Changes section should exist');
    assert.notEqual(changedIndex, -1, 'Changed section should exist');
    assert.notEqual(fixedIndex, -1, 'Fixed section should exist');
    assert.ok(breakingIndex < changedIndex, 'Breaking Changes should appear before Changed');
    assert.ok(changedIndex < fixedIndex, 'Changed should appear before Fixed');

    const changelogCall = stub.calls.find((call) => modeFromPrompt(call.input) === 'changelog');
    assert.ok(changelogCall, 'expected at least one changelog-mode Claude invocation');
    assert.match(
      changelogCall.input,
      /breaking: true/,
      'breaking metadata should reach the prompt verbatim',
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

test('writePrereleaseNotesForVersion writes cumulative beta notes without mutating stable changelog artifacts', async () => {
  const { writePrereleaseNotesForVersion } = await loadModule();
  const workspace = createWorkspace('prerelease-beta-notes');
  const projectRoot = path.join(workspace, 'SubMiner');
  const changelogPath = path.join(projectRoot, 'CHANGELOG.md');
  const docsChangelogPath = path.join(projectRoot, 'docs-site', 'changelog.md');
  const existingChangelog = '# Changelog\n\n## v0.11.2 (2026-04-07)\n- Stable release.\n';
  const existingDocsChangelog = '# Changelog\n\n## v0.11.2 (2026-04-07)\n- Stable docs release.\n';

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.mkdirSync(path.join(projectRoot, 'docs-site'), { recursive: true });
  fs.writeFileSync(
    path.join(projectRoot, 'package.json'),
    JSON.stringify({ name: 'subminer', version: '0.11.3-beta.1' }, null, 2),
    'utf8',
  );
  fs.writeFileSync(changelogPath, existingChangelog, 'utf8');
  fs.writeFileSync(docsChangelogPath, existingDocsChangelog, 'utf8');
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: added', 'area: overlay', '', '- Added prerelease coverage.'].join('\n'),
    'utf8',
  );
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '002.md'),
    ['type: fixed', 'area: launcher', '', '- Fixed prerelease packaging checks.'].join('\n'),
    'utf8',
  );

  try {
    const stub = defaultStubClaude();
    const outputPath = writePrereleaseNotesForVersion({
      cwd: projectRoot,
      version: '0.11.3-beta.1',
      deps: { runClaude: stub.runClaude },
    });

    assert.equal(outputPath, path.join(projectRoot, 'release', 'prerelease-notes.md'));
    assert.equal(
      fs.readFileSync(changelogPath, 'utf8'),
      existingChangelog,
      'stable CHANGELOG.md should remain unchanged',
    );
    assert.equal(
      fs.readFileSync(docsChangelogPath, 'utf8'),
      existingDocsChangelog,
      'docs-site changelog should remain unchanged',
    );
    assert.equal(fs.existsSync(path.join(projectRoot, 'changes', '001.md')), true);
    assert.equal(fs.existsSync(path.join(projectRoot, 'changes', '002.md')), true);

    assert.equal(stub.calls.length, 1, 'prerelease should issue exactly one Claude call');
    assert.equal(modeFromPrompt(stub.calls[0]!.input), 'release-notes');

    const prereleaseNotes = fs.readFileSync(outputPath, 'utf8');
    assert.match(prereleaseNotes, /^> This is a prerelease build for testing\./m);
    assert.match(prereleaseNotes, /## Highlights\n### Added\n- Polished: added entry\./);
    assert.match(prereleaseNotes, /### Fixed\n- Polished: fixed entry\./);
    assert.match(prereleaseNotes, /## Installation\n\nSee the README and docs\/installation guide/);
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('writePrereleaseNotesForVersion supports rc prereleases', async () => {
  const { writePrereleaseNotesForVersion } = await loadModule();
  const workspace = createWorkspace('prerelease-rc-notes');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(
    path.join(projectRoot, 'package.json'),
    JSON.stringify({ name: 'subminer', version: '0.11.3-rc.1' }, null, 2),
    'utf8',
  );
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: changed', 'area: release', '', '- Prepared release candidate notes.'].join('\n'),
    'utf8',
  );

  try {
    const stub = defaultStubClaude();
    const outputPath = writePrereleaseNotesForVersion({
      cwd: projectRoot,
      version: '0.11.3-rc.1',
      deps: { runClaude: stub.runClaude },
    });

    const prereleaseNotes = fs.readFileSync(outputPath, 'utf8');
    assert.match(prereleaseNotes, /## Highlights\n### Changed\n- Polished: changed entry\./);
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('writePrereleaseNotesForVersion rejects unsupported prerelease identifiers', async () => {
  const { writePrereleaseNotesForVersion } = await loadModule();
  const workspace = createWorkspace('prerelease-alpha-reject');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(
    path.join(projectRoot, 'package.json'),
    JSON.stringify({ name: 'subminer', version: '0.11.3-alpha.1' }, null, 2),
    'utf8',
  );
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: added', 'area: overlay', '', '- Unsupported alpha prerelease.'].join('\n'),
    'utf8',
  );

  try {
    assert.throws(
      () =>
        writePrereleaseNotesForVersion({
          cwd: projectRoot,
          version: '0.11.3-alpha.1',
        }),
      /Unsupported prerelease version \(0\.11\.3-alpha\.1\)/,
    );
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('writePrereleaseNotesForVersion rejects mismatched package versions', async () => {
  const { writePrereleaseNotesForVersion } = await loadModule();
  const workspace = createWorkspace('prerelease-version-mismatch');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(
    path.join(projectRoot, 'package.json'),
    JSON.stringify({ name: 'subminer', version: '0.11.3-beta.1' }, null, 2),
    'utf8',
  );
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: added', 'area: overlay', '', '- Mismatched prerelease.'].join('\n'),
    'utf8',
  );

  try {
    assert.throws(
      () =>
        writePrereleaseNotesForVersion({
          cwd: projectRoot,
          version: '0.11.3-beta.2',
        }),
      /package\.json version \(0\.11\.3-beta\.1\) does not match requested release version \(0\.11\.3-beta\.2\)/,
    );
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('writePrereleaseNotesForVersion rejects empty prerelease note generation when no fragments exist', async () => {
  const { writePrereleaseNotesForVersion } = await loadModule();
  const workspace = createWorkspace('prerelease-no-fragments');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(
    path.join(projectRoot, 'package.json'),
    JSON.stringify({ name: 'subminer', version: '0.11.3-beta.1' }, null, 2),
    'utf8',
  );

  try {
    assert.throws(
      () =>
        writePrereleaseNotesForVersion({
          cwd: projectRoot,
          version: '0.11.3-beta.1',
        }),
      /No changelog fragments found in changes\//,
    );
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('writeChangelogArtifacts surfaces a clear error when claude is missing from PATH', async () => {
  const { writeChangelogArtifacts } = await loadModule();
  const workspace = createWorkspace('claude-missing');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(path.join(projectRoot, 'CHANGELOG.md'), '# Changelog\n', 'utf8');
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: added', 'area: overlay', '', '- A change.'].join('\n'),
    'utf8',
  );

  // The production defaultRunClaude wrapper translates ENOENT into this friendly
  // message; we simulate that contract here so the test exercises the propagation
  // path through polishFragmentsWithClaude rather than re-implementing the
  // execFileSync mock.
  const enoent = (): string => {
    throw new Error(
      "claude CLI not found on PATH. Install Claude Code (https://claude.com/claude-code) and ensure 'claude' is on your PATH before running changelog:build.",
    );
  };

  try {
    assert.throws(
      () =>
        writeChangelogArtifacts({
          cwd: projectRoot,
          version: '0.5.0',
          date: '2026-04-06',
          deps: { runClaude: enoent },
        }),
      /claude CLI not found on PATH/,
    );
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('writeChangelogArtifacts rejects empty claude output', async () => {
  const { writeChangelogArtifacts } = await loadModule();
  const workspace = createWorkspace('claude-empty');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(path.join(projectRoot, 'CHANGELOG.md'), '# Changelog\n', 'utf8');
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: added', 'area: overlay', '', '- A change.'].join('\n'),
    'utf8',
  );

  try {
    assert.throws(
      () =>
        writeChangelogArtifacts({
          cwd: projectRoot,
          version: '0.5.0',
          date: '2026-04-06',
          deps: { runClaude: () => '   \n  ' },
        }),
      /claude returned empty output/,
    );
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('writeChangelogArtifacts rejects claude output missing required section headers', async () => {
  const { writeChangelogArtifacts } = await loadModule();
  const workspace = createWorkspace('claude-no-headers');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(path.join(projectRoot, 'CHANGELOG.md'), '# Changelog\n', 'utf8');
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: added', 'area: overlay', '', '- A change.'].join('\n'),
    'utf8',
  );

  try {
    assert.throws(
      () =>
        writeChangelogArtifacts({
          cwd: projectRoot,
          version: '0.5.0',
          date: '2026-04-06',
          deps: { runClaude: () => 'Sure, here is your changelog: it is great.' },
        }),
      /missing the expected section heading/,
    );
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('writeChangelogArtifacts rejects changelog-mode output that omits the Internal <details> wrapper when internal fragments are present', async () => {
  const { writeChangelogArtifacts } = await loadModule();
  const workspace = createWorkspace('claude-no-details');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(path.join(projectRoot, 'CHANGELOG.md'), '# Changelog\n', 'utf8');
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: added', 'area: overlay', '', '- A user-facing change.'].join('\n'),
    'utf8',
  );
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '002.md'),
    ['type: internal', 'area: release', '', '- An internal note.'].join('\n'),
    'utf8',
  );

  const noDetailsResponder = (input: string): string => {
    if (modeFromPrompt(input) === 'changelog') {
      return '### Added\n- Polished: added.\n\n### Internal\n- Polished: internal (no details wrapper).';
    }
    return defaultPolishedBody(input);
  };

  try {
    assert.throws(
      () =>
        writeChangelogArtifacts({
          cwd: projectRoot,
          version: '0.5.0',
          date: '2026-04-06',
          deps: { runClaude: noDetailsResponder },
        }),
      /<details><summary>Internal changes<\/summary> wrapper/,
    );
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('writeChangelogArtifacts filters internal fragments from the release-notes Claude prompt', async () => {
  const { writeChangelogArtifacts } = await loadModule();
  const workspace = createWorkspace('release-notes-internal-filter');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(path.join(projectRoot, 'CHANGELOG.md'), '# Changelog\n', 'utf8');
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: added', 'area: overlay', '', '- A user-facing change.'].join('\n'),
    'utf8',
  );
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '002.md'),
    ['type: internal', 'area: release', '', '- An internal CI tweak.'].join('\n'),
    'utf8',
  );

  try {
    const stub = defaultStubClaude();
    writeChangelogArtifacts({
      cwd: projectRoot,
      version: '0.5.0',
      date: '2026-04-06',
      deps: { runClaude: stub.runClaude },
    });

    const changelogCall = stub.calls.find((call) => modeFromPrompt(call.input) === 'changelog');
    const releaseNotesCall = stub.calls.find(
      (call) => modeFromPrompt(call.input) === 'release-notes',
    );
    assert.ok(changelogCall, 'expected a changelog-mode invocation');
    assert.ok(releaseNotesCall, 'expected a release-notes-mode invocation');

    assert.deepEqual(
      fragmentTypesInPrompt(changelogCall.input).sort(),
      ['added', 'internal'],
      'changelog mode keeps internal fragments',
    );
    assert.deepEqual(
      fragmentTypesInPrompt(releaseNotesCall.input),
      ['added'],
      'release-notes mode drops internal fragments',
    );

    const releaseNotes = fs.readFileSync(
      path.join(projectRoot, 'release', 'release-notes.md'),
      'utf8',
    );
    assert.doesNotMatch(releaseNotes, /<details>/);
    assert.doesNotMatch(releaseNotes, /### Internal/);

    const changelog = fs.readFileSync(path.join(projectRoot, 'CHANGELOG.md'), 'utf8');
    assert.match(changelog, /<details>[\s\S]*<summary>Internal changes<\/summary>/);
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('writeChangelogArtifacts strips <details> blocks from release notes when reusing an existing CHANGELOG section', async () => {
  const { writeChangelogArtifacts } = await loadModule();
  const workspace = createWorkspace('reuse-existing-section');
  const projectRoot = path.join(workspace, 'SubMiner');
  const existingChangelog = [
    '# Changelog',
    '',
    '## v0.4.1 (2026-03-07)',
    '### Added',
    '- Polished: previously committed.',
    '',
    '<details>',
    '<summary>Internal changes</summary>',
    '',
    '### Internal',
    '- Polished: internal note.',
    '',
    '</details>',
    '',
  ].join('\n');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(path.join(projectRoot, 'CHANGELOG.md'), existingChangelog, 'utf8');
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: added', 'area: overlay', '', '- Stale fragment.'].join('\n'),
    'utf8',
  );

  try {
    const stub = defaultStubClaude();
    writeChangelogArtifacts({
      cwd: projectRoot,
      version: '0.4.1',
      date: '2026-03-08',
      deps: { runClaude: stub.runClaude },
    });

    assert.equal(
      stub.calls.length,
      0,
      'no Claude calls should fire when the section already exists',
    );

    const releaseNotes = fs.readFileSync(
      path.join(projectRoot, 'release', 'release-notes.md'),
      'utf8',
    );
    assert.match(releaseNotes, /## Highlights\n### Added\n- Polished: previously committed\./);
    assert.doesNotMatch(releaseNotes, /<details>/);
    assert.doesNotMatch(releaseNotes, /### Internal/);
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

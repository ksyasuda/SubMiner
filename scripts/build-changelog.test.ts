import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
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

function assertPromptRequestsNestedBullets(input: string): void {
  assert.match(input, /In both modes, split every item into one nested bullet per distinct change/);
  assert.match(input, /Never stack several distinct changes into one long paragraph-shaped bullet/);
  assert.match(input, /Keep nested bullets short, concrete, and readable by non-technical users/);
  assert.match(input, /Avoid paragraph-style release-note bullets/);
}

function assertReleaseNotesPromptRequestsNestedBullets(input: string): void {
  assertPromptRequestsNestedBullets(input);
  assert.match(
    input,
    /In MODE: release-notes, nested bullets should also cover user benefit and any user action/,
  );
  assert.match(input, /Do not require the exact nested labels/);
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
    assert.match(releaseNotes, /- Windows: `SubMiner-\*\.exe` and `SubMiner-\*-win\.zip`/);
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

test('writeChangelogArtifacts prompts Claude to summarize the final stable outcome instead of prerelease churn', async () => {
  const { writeChangelogArtifacts } = await loadModule();
  const workspace = createWorkspace('stable-outcome-prompt');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(projectRoot, { recursive: true });
  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(path.join(projectRoot, 'CHANGELOG.md'), '# Changelog\n', 'utf8');
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    [
      'type: added',
      'area: config',
      '',
      '- Added a dedicated Config window with launcher entry points.',
    ].join('\n'),
    'utf8',
  );
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '002.md'),
    [
      'type: changed',
      'area: config',
      'breaking: true',
      '',
      '- Renamed the Config window to Settings window and changed the launcher entry point to `subminer settings`.',
    ].join('\n'),
    'utf8',
  );
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '003.md'),
    [
      'type: fixed',
      'area: config',
      '',
      '- Fixed Settings window search and live subtitle CSS saves.',
    ].join('\n'),
    'utf8',
  );

  try {
    const stub = defaultStubClaude();
    writeChangelogArtifacts({
      cwd: projectRoot,
      version: '0.12.0',
      date: '2026-05-24',
      deps: { runClaude: stub.runClaude },
    });

    const prompts = stub.calls.map((call) => call.input);
    assert.equal(prompts.length, 2, 'expected changelog and release-notes prompts');
    for (const prompt of prompts) {
      assert.match(prompt, /Treat the fragment list as one cumulative release outcome/);
      assert.match(
        prompt,
        /only if the final release requires action from users upgrading from the previous stable release/,
      );
      assert.match(prompt, /Config window.*Settings window/s);
      assert.match(
        prompt,
        /Multiple fixes within the same prerelease cycle should collapse into one current-state bullet/,
      );
      assertPromptRequestsNestedBullets(prompt);
    }

    const releaseNotesPrompt = stub.calls.find(
      (call) => modeFromPrompt(call.input) === 'release-notes',
    );
    assert.ok(releaseNotesPrompt, 'expected a release-notes Claude invocation');
    assertReleaseNotesPromptRequestsNestedBullets(releaseNotesPrompt.input);
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
    /requires a reconciled changelog fragment/,
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
    /requires a reconciled changelog fragment/,
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

  assert.doesNotThrow(() =>
    verifyPullRequestChangelog({
      changedEntries: [
        { path: 'src/main-entry.ts', status: 'M' },
        { path: 'changes/001.md', status: 'M' },
      ],
      changedLabels: [],
    }),
  );

  assert.doesNotThrow(() =>
    verifyPullRequestChangelog({
      changedEntries: [
        { path: 'src/main-entry.ts', status: 'M' },
        { path: 'changes/001.md', status: 'A' },
        { path: 'changes/002.md', status: 'A' },
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
      deps: { runClaude: stub.runClaude, listPrereleaseTags: () => [] },
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
    assert.match(prereleaseNotes, /<!-- prerelease-version: 0\.11\.3-beta\.1 -->/);
    assert.doesNotMatch(prereleaseNotes, /## Changes since /);
    assert.match(prereleaseNotes, /## Highlights\n### Added\n- Polished: added entry\./);
    assert.match(prereleaseNotes, /### Fixed\n- Polished: fixed entry\./);
    assert.match(prereleaseNotes, /## Installation\n\nSee the README and docs\/installation guide/);
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('writePrereleaseNotesForVersion reuses existing prerelease notes when adding new fragments', async () => {
  const { writePrereleaseNotesForVersion } = await loadModule();
  const workspace = createWorkspace('prerelease-reuse-existing-notes');
  const projectRoot = path.join(workspace, 'SubMiner');
  const existingNotes = [
    '> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.',
    '',
    '<!-- prerelease-base-version: 0.11.3 -->',
    '',
    '## Highlights',
    '### Added',
    '- Overlay: Previous beta entry.',
    '',
    '## Installation',
    '',
    'See the README and docs/installation guide for full setup steps.',
    '',
    '## Assets',
    '',
    '- Linux: `SubMiner.AppImage`',
    '',
  ].join('\n');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.mkdirSync(path.join(projectRoot, 'release'), { recursive: true });
  fs.writeFileSync(
    path.join(projectRoot, 'package.json'),
    JSON.stringify({ name: 'subminer', version: '0.11.3-beta.2' }, null, 2),
    'utf8',
  );
  fs.writeFileSync(path.join(projectRoot, 'release', 'prerelease-notes.md'), existingNotes, 'utf8');
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: fixed', 'area: launcher', '', '- Fixed launcher prerelease packaging.'].join('\n'),
    'utf8',
  );

  try {
    const stub = recordingRunClaude((input) => {
      if (!input.includes('Overlay: Previous beta entry.')) {
        return '### Fixed\n- Launcher: Added only the latest fix.';
      }
      return [
        '### Added',
        '- Overlay: Previous beta entry.',
        '',
        '### Fixed',
        '- Launcher: Added only the latest fix.',
      ].join('\n');
    });

    const outputPath = writePrereleaseNotesForVersion({
      cwd: projectRoot,
      version: '0.11.3-beta.2',
      deps: { runClaude: stub.runClaude, listPrereleaseTags: () => [] },
    });

    assert.equal(stub.calls.length, 1, 'prerelease should issue exactly one Claude call');
    assert.match(stub.calls[0]!.input, /EXISTING PRERELEASE NOTES/);

    const prereleaseNotes = fs.readFileSync(outputPath, 'utf8');
    assert.match(prereleaseNotes, /- Overlay: Previous beta entry\./);
    assert.match(prereleaseNotes, /- Launcher: Added only the latest fix\./);
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('writePrereleaseNotesForVersion ignores unmarked prerelease notes from an older release line', async () => {
  const { writePrereleaseNotesForVersion } = await loadModule();
  const workspace = createWorkspace('prerelease-ignore-unmarked-old-notes');
  const projectRoot = path.join(workspace, 'SubMiner');
  const existingNotes = [
    '> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.',
    '',
    '## Highlights',
    '### Added',
    '- Settings Window: Previous release line entry.',
    '',
    '## Installation',
    '',
    'See the README and docs/installation guide for full setup steps.',
    '',
  ].join('\n');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.mkdirSync(path.join(projectRoot, 'release'), { recursive: true });
  fs.writeFileSync(
    path.join(projectRoot, 'package.json'),
    JSON.stringify({ name: 'subminer', version: '0.17.0-beta.1' }, null, 2),
    'utf8',
  );
  fs.writeFileSync(path.join(projectRoot, 'release', 'prerelease-notes.md'), existingNotes, 'utf8');
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    [
      'type: changed',
      'area: overlay',
      '',
      '- Replaced subtitle delay actions with native mpv keybindings.',
    ].join('\n'),
    'utf8',
  );

  try {
    const stub = defaultStubClaude();
    const outputPath = writePrereleaseNotesForVersion({
      cwd: projectRoot,
      version: '0.17.0-beta.1',
      deps: { runClaude: stub.runClaude, listPrereleaseTags: () => [] },
    });

    assert.equal(stub.calls.length, 1, 'prerelease should issue exactly one Claude call');
    assert.doesNotMatch(stub.calls[0]!.input, /EXISTING PRERELEASE NOTES/);
    assert.doesNotMatch(stub.calls[0]!.input, /Settings Window: Previous release line entry/);

    const prereleaseNotes = fs.readFileSync(outputPath, 'utf8');
    assert.match(prereleaseNotes, /### Changed\n- Polished: changed entry\./);
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('writePrereleaseNotesForVersion prompts Claude to revise stale prerelease bullets instead of appending fix churn', async () => {
  const { writePrereleaseNotesForVersion } = await loadModule();
  const workspace = createWorkspace('prerelease-net-outcome-prompt');
  const projectRoot = path.join(workspace, 'SubMiner');
  const existingNotes = [
    '> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.',
    '',
    '<!-- prerelease-base-version: 0.12.0 -->',
    '',
    '## Highlights',
    '### Added',
    '- Config Window: Previous beta entry.',
    '',
    '## Installation',
    '',
    'See the README and docs/installation guide for full setup steps.',
    '',
    '## Assets',
    '',
    '- Linux: `SubMiner.AppImage`',
    '',
  ].join('\n');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.mkdirSync(path.join(projectRoot, 'release'), { recursive: true });
  fs.writeFileSync(
    path.join(projectRoot, 'package.json'),
    JSON.stringify({ name: 'subminer', version: '0.12.0-beta.2' }, null, 2),
    'utf8',
  );
  fs.writeFileSync(path.join(projectRoot, 'release', 'prerelease-notes.md'), existingNotes, 'utf8');
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    [
      'type: changed',
      'area: config',
      'breaking: true',
      '',
      '- Renamed the Config window to Settings window.',
    ].join('\n'),
    'utf8',
  );
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '002.md'),
    ['type: fixed', 'area: config', '', '- Fixed Settings window search.'].join('\n'),
    'utf8',
  );

  try {
    const stub = recordingRunClaude(() => '### Added\n- Settings Window: Current beta state.');
    writePrereleaseNotesForVersion({
      cwd: projectRoot,
      version: '0.12.0-beta.2',
      deps: { runClaude: stub.runClaude, listPrereleaseTags: () => [] },
    });

    assert.equal(stub.calls.length, 1, 'prerelease should issue exactly one Claude call');
    const prompt = stub.calls[0]!.input;
    assert.match(prompt, /EXISTING PRERELEASE NOTES/);
    assert.match(prompt, /Existing prerelease notes are a baseline, not an immutable changelog/);
    assert.match(prompt, /replace stale beta or RC wording/);
    assert.match(
      prompt,
      /Multiple fixes within the same prerelease cycle should collapse into one current-state bullet/,
    );
    assertReleaseNotesPromptRequestsNestedBullets(prompt);
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
      deps: { runClaude: stub.runClaude, listPrereleaseTags: () => [] },
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

test('writeChangelogArtifacts appends contributor attribution and a new-contributors section to release notes', async () => {
  const { writeChangelogArtifacts } = await loadModule();
  const workspace = createWorkspace('release-notes-contributors');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(path.join(projectRoot, 'CHANGELOG.md'), '# Changelog\n', 'utf8');
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: added', 'area: overlay', '', '- Added a feature.'].join('\n'),
    'utf8',
  );
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '002.md'),
    ['type: fixed', 'area: jellyfin', '', '- Fixed a bug.'].join('\n'),
    'utf8',
  );

  try {
    const stub = defaultStubClaude();
    const resolveContributionsCalls: string[][] = [];
    writeChangelogArtifacts({
      cwd: projectRoot,
      version: '0.6.0',
      date: '2026-05-06',
      deps: {
        runClaude: stub.runClaude,
        resolveContributions: (fragmentPaths) => {
          resolveContributionsCalls.push(fragmentPaths);
          return [
            {
              prNumber: 110,
              login: 'ksyasuda',
              title: 'feat(overlay): add a feature',
              isFirstContribution: false,
            },
            {
              prNumber: 112,
              login: 'bee-san',
              title: 'fix(jellyfin): restart remote session',
              isFirstContribution: true,
            },
          ];
        },
      },
    });

    assert.equal(resolveContributionsCalls.length, 1, 'resolves contributions once per release');
    assert.deepEqual(resolveContributionsCalls[0], [
      path.join(projectRoot, 'changes', '001.md'),
      path.join(projectRoot, 'changes', '002.md'),
    ]);

    const releaseNotes = fs.readFileSync(
      path.join(projectRoot, 'release', 'release-notes.md'),
      'utf8',
    );
    assert.match(releaseNotes, /## What's Changed\n\n/);
    assert.match(releaseNotes, /- feat\(overlay\): add a feature by @ksyasuda in #110\n/);
    assert.match(releaseNotes, /- fix\(jellyfin\): restart remote session by @bee-san in #112\n/);
    assert.match(
      releaseNotes,
      /## New Contributors\n\n- @bee-san made their first contribution in #112/,
    );
    assert.ok(
      releaseNotes.indexOf("## What's Changed") > releaseNotes.indexOf('## Highlights'),
      "What's Changed should follow Highlights",
    );
    assert.ok(
      releaseNotes.indexOf('## New Contributors') < releaseNotes.indexOf('## Installation'),
      'contributor attribution should appear before Installation',
    );
    assert.doesNotMatch(releaseNotes, /## What’s Changed/);
    assert.doesNotMatch(
      releaseNotes,
      /ksyasuda made their first contribution/,
      'returning contributors are not listed under New Contributors',
    );

    // Attribution is a release-notes concern only; the CHANGELOG stays clean.
    const changelog = fs.readFileSync(path.join(projectRoot, 'CHANGELOG.md'), 'utf8');
    assert.doesNotMatch(changelog, /What's Changed|What’s Changed/);
    assert.doesNotMatch(changelog, /New Contributors/);
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('writeChangelogArtifacts skips contributor attribution in GitHub Actions without a token', async () => {
  const { writeChangelogArtifacts } = await loadModule();
  const workspace = createWorkspace('release-notes-actions-no-token');
  const projectRoot = path.join(workspace, 'SubMiner');
  const originalActions = process.env.GITHUB_ACTIONS;
  const originalGhToken = process.env.GH_TOKEN;
  const originalGithubToken = process.env.GITHUB_TOKEN;
  const originalPath = process.env.PATH;
  const originalWarn = console.warn;
  const warnings: string[] = [];

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(path.join(projectRoot, 'CHANGELOG.md'), '# Changelog\n', 'utf8');
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: added', 'area: release', '', '- Added a feature.'].join('\n'),
    'utf8',
  );

  try {
    process.env.GITHUB_ACTIONS = 'true';
    delete process.env.GH_TOKEN;
    delete process.env.GITHUB_TOKEN;
    process.env.PATH = workspace;
    console.warn = (message?: unknown) => {
      warnings.push(String(message));
    };

    writeChangelogArtifacts({
      cwd: projectRoot,
      version: '0.6.0',
      date: '2026-05-06',
      deps: { runClaude: defaultStubClaude().runClaude },
    });

    assert.deepEqual(warnings, []);
    const releaseNotes = fs.readFileSync(
      path.join(projectRoot, 'release', 'release-notes.md'),
      'utf8',
    );
    assert.doesNotMatch(releaseNotes, /## What's Changed/);
  } finally {
    console.warn = originalWarn;
    if (originalActions === undefined) {
      delete process.env.GITHUB_ACTIONS;
    } else {
      process.env.GITHUB_ACTIONS = originalActions;
    }
    if (originalGhToken === undefined) {
      delete process.env.GH_TOKEN;
    } else {
      process.env.GH_TOKEN = originalGhToken;
    }
    if (originalGithubToken === undefined) {
      delete process.env.GITHUB_TOKEN;
    } else {
      process.env.GITHUB_TOKEN = originalGithubToken;
    }
    if (originalPath === undefined) {
      delete process.env.PATH;
    } else {
      process.env.PATH = originalPath;
    }
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('shouldSkipDefaultContributionLookup skips GitHub Actions without a gh token', async () => {
  const { shouldSkipDefaultContributionLookup } = await loadModule();

  assert.equal(
    shouldSkipDefaultContributionLookup({
      GITHUB_ACTIONS: 'true',
      GH_TOKEN: undefined,
      GITHUB_TOKEN: undefined,
    }),
    true,
  );
  assert.equal(
    shouldSkipDefaultContributionLookup({
      GITHUB_ACTIONS: 'true',
      GH_TOKEN: 'ghs_test',
      GITHUB_TOKEN: undefined,
    }),
    false,
  );
  assert.equal(
    shouldSkipDefaultContributionLookup({
      GITHUB_ACTIONS: undefined,
      GH_TOKEN: undefined,
      GITHUB_TOKEN: undefined,
    }),
    false,
  );
});

test('writeReleaseNotesForVersion preserves committed contributor attribution before installation', async () => {
  const { writeReleaseNotesForVersion } = await loadModule();
  const workspace = createWorkspace('release-notes-preserve-attribution');
  const projectRoot = path.join(workspace, 'SubMiner');
  const existingChangelog = [
    '# Changelog',
    '',
    '## v0.8.0 (2026-04-17)',
    '### Added',
    '- Polished: released feature.',
    '',
    '<details>',
    '<summary>Internal changes</summary>',
    '',
    '### Internal',
    '- Polished: internal release note.',
    '',
    '</details>',
    '',
  ].join('\n');
  const committedReleaseNotes = [
    '## Highlights',
    '### Added',
    '- Old generated body.',
    '',
    '## Installation',
    '',
    'See the README and docs/installation guide for full setup steps.',
    '',
    '## Assets',
    '',
    '- Linux: `SubMiner.AppImage`',
    '',
    '## What’s Changed',
    '',
    '- feat(release): add contributor attribution by @ksyasuda in #114',
    '',
    '## New Contributors',
    '',
    '- @bee-san made their first contribution in #112',
    '',
  ].join('\n');

  fs.mkdirSync(path.join(projectRoot, 'release'), { recursive: true });
  fs.writeFileSync(path.join(projectRoot, 'CHANGELOG.md'), existingChangelog, 'utf8');
  fs.writeFileSync(
    path.join(projectRoot, 'release', 'release-notes.md'),
    committedReleaseNotes,
    'utf8',
  );

  try {
    const outputPath = writeReleaseNotesForVersion({
      cwd: projectRoot,
      version: '0.8.0',
    });
    const releaseNotes = fs.readFileSync(outputPath, 'utf8');

    assert.match(releaseNotes, /## Highlights\n### Added\n- Polished: released feature\./);
    assert.doesNotMatch(releaseNotes, /<details>/);
    assert.doesNotMatch(releaseNotes, /### Internal/);
    assert.match(
      releaseNotes,
      /## What's Changed\n\n- feat\(release\): add contributor attribution by @ksyasuda in #114/,
    );
    assert.match(
      releaseNotes,
      /## New Contributors\n\n- @bee-san made their first contribution in #112/,
    );
    assert.ok(
      releaseNotes.indexOf("## What's Changed") > releaseNotes.indexOf('## Highlights'),
      "What's Changed should follow Highlights",
    );
    assert.ok(
      releaseNotes.indexOf('## New Contributors') < releaseNotes.indexOf('## Installation'),
      'New Contributors should appear before Installation',
    );
    assert.doesNotMatch(releaseNotes, /## What’s Changed/);
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

test('selectPreviousPrereleaseTag orders betas before rcs and filters other base versions', async () => {
  const { selectPreviousPrereleaseTag } = await loadModule();

  const tags = [
    'v0.19.4-beta.1',
    'v0.19.4-beta.3',
    'v0.19.4-beta.2',
    'v0.19.3-beta.9',
    'v0.19.4-rc.1',
    'not-a-tag',
  ];

  assert.equal(selectPreviousPrereleaseTag(tags, '0.19.4-beta.1'), null);
  assert.equal(selectPreviousPrereleaseTag(tags, '0.19.4-beta.2'), 'v0.19.4-beta.1');
  assert.equal(selectPreviousPrereleaseTag(tags, '0.19.4-beta.4'), 'v0.19.4-beta.3');
  assert.equal(selectPreviousPrereleaseTag(tags, '0.19.4-rc.1'), 'v0.19.4-beta.3');
  assert.equal(selectPreviousPrereleaseTag(tags, '0.19.4-rc.2'), 'v0.19.4-rc.1');
  // Regenerating notes for an already-tagged version must not pick itself.
  assert.equal(selectPreviousPrereleaseTag(tags, '0.19.4-beta.3'), 'v0.19.4-beta.2');
  assert.equal(selectPreviousPrereleaseTag(['v0.19.3-beta.1'], '0.19.4-beta.2'), null);
});

test('writePrereleaseNotesForVersion adds a delta section generated from fragment diffs', async () => {
  const { writePrereleaseNotesForVersion } = await loadModule();
  const workspace = createWorkspace('prerelease-delta-section');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(
    path.join(projectRoot, 'package.json'),
    JSON.stringify({ name: 'subminer', version: '0.12.0-beta.2' }, null, 2),
    'utf8',
  );
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: fixed', 'area: overlay', '', '- Fixed overlay focus and macOS helper.'].join('\n'),
    'utf8',
  );

  try {
    const stub = recordingRunClaude((input) =>
      input.includes('MODIFIED FRAGMENT')
        ? '- Fixed the macOS helper deployment target for older systems.'
        : '### Fixed\n- Overlay: cumulative fixed entry.',
    );
    const outputPath = writePrereleaseNotesForVersion({
      cwd: projectRoot,
      version: '0.12.0-beta.2',
      deps: {
        runClaude: stub.runClaude,
        listPrereleaseTags: () => ['v0.12.0-beta.1'],
        resolveFragmentDelta: (_cwd, previousTag) => {
          assert.equal(previousTag, 'v0.12.0-beta.1');
          return [
            {
              path: 'changes/002.md',
              status: 'added',
              after: 'type: fixed\narea: macos\n\n- Fixed helper deployment target.',
            },
            {
              path: 'changes/001.md',
              status: 'modified',
              before: '- Fixed overlay focus.',
              after: '- Fixed overlay focus and macOS helper.',
            },
            {
              path: 'changes/003.md',
              status: 'deleted',
              before: 'type: added\narea: stats\n\n- Reverted experimental stats view.',
            },
          ];
        },
      },
    });

    assert.equal(stub.calls.length, 2, 'delta and cumulative polish are separate Claude calls');
    const deltaPrompt = stub.calls[0]!.input;
    assert.match(deltaPrompt, /ADDED FRAGMENT changes\/002\.md/);
    assert.match(deltaPrompt, /MODIFIED FRAGMENT changes\/001\.md/);
    assert.match(deltaPrompt, /BEFORE:\n- Fixed overlay focus\./);
    assert.match(deltaPrompt, /AFTER:\n- Fixed overlay focus and macOS helper\./);
    assert.match(deltaPrompt, /DELETED FRAGMENT changes\/003\.md/);
    assert.match(deltaPrompt, /If the edit is editorial/);
    assert.match(deltaPrompt, /removed or reverted/);
    assert.match(deltaPrompt, /No user-facing changes since v0\.12\.0-beta\.1\./);
    assert.equal(modeFromPrompt(stub.calls[1]!.input), 'release-notes');

    const prereleaseNotes = fs.readFileSync(outputPath, 'utf8');
    assert.match(
      prereleaseNotes,
      /<!-- prerelease-version: 0\.12\.0-beta\.2; since: v0\.12\.0-beta\.1 -->/,
    );
    const deltaIndex = prereleaseNotes.indexOf('## Changes since v0.12.0-beta.1');
    const highlightsIndex = prereleaseNotes.indexOf('## Highlights');
    assert.ok(deltaIndex !== -1, 'delta section heading should be present');
    assert.ok(deltaIndex < highlightsIndex, 'delta section should precede Highlights');
    assert.match(prereleaseNotes, /- Fixed the macOS helper deployment target for older systems\./);
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('writePrereleaseNotesForVersion renders a fallback delta line when no fragments changed', async () => {
  const { writePrereleaseNotesForVersion } = await loadModule();
  const workspace = createWorkspace('prerelease-empty-delta');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(
    path.join(projectRoot, 'package.json'),
    JSON.stringify({ name: 'subminer', version: '0.12.0-beta.3' }, null, 2),
    'utf8',
  );
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: fixed', 'area: overlay', '', '- Fixed overlay focus.'].join('\n'),
    'utf8',
  );

  try {
    const stub = defaultStubClaude();
    const outputPath = writePrereleaseNotesForVersion({
      cwd: projectRoot,
      version: '0.12.0-beta.3',
      deps: {
        runClaude: stub.runClaude,
        listPrereleaseTags: () => ['v0.12.0-beta.1', 'v0.12.0-beta.2'],
        resolveFragmentDelta: () => [],
      },
    });

    assert.equal(stub.calls.length, 1, 'empty delta must not spend a Claude call');
    const prereleaseNotes = fs.readFileSync(outputPath, 'utf8');
    assert.match(
      prereleaseNotes,
      /## Changes since v0\.12\.0-beta\.2\n\n- No changelog fragment changes since v0\.12\.0-beta\.2; this build contains packaging or internal-only updates\./,
    );
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('writePrereleaseNotesForVersion rejects non-bullet delta output from Claude', async () => {
  const { writePrereleaseNotesForVersion } = await loadModule();
  const workspace = createWorkspace('prerelease-delta-invalid-output');
  const projectRoot = path.join(workspace, 'SubMiner');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(
    path.join(projectRoot, 'package.json'),
    JSON.stringify({ name: 'subminer', version: '0.12.0-beta.2' }, null, 2),
    'utf8',
  );
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: fixed', 'area: overlay', '', '- Fixed overlay focus.'].join('\n'),
    'utf8',
  );

  try {
    const stub = recordingRunClaude(() => 'Here are the changes:\n- One change.');
    assert.throws(
      () =>
        writePrereleaseNotesForVersion({
          cwd: projectRoot,
          version: '0.12.0-beta.2',
          deps: {
            runClaude: stub.runClaude,
            listPrereleaseTags: () => ['v0.12.0-beta.1'],
            resolveFragmentDelta: () => [
              { path: 'changes/001.md', status: 'added', after: '- Fixed overlay focus.' },
            ],
          },
        }),
      /delta output must contain only Markdown bullets/,
    );
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('writePrereleaseNotesForVersion strips the stale delta section from the reused baseline', async () => {
  const { writePrereleaseNotesForVersion } = await loadModule();
  const workspace = createWorkspace('prerelease-reuse-strips-delta');
  const projectRoot = path.join(workspace, 'SubMiner');
  const existingNotes = [
    '> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.',
    '',
    '<!-- prerelease-version: 0.12.0-beta.2; since: v0.12.0-beta.1 -->',
    '',
    '## Changes since v0.12.0-beta.1',
    '',
    '- Stale beta-to-beta delta bullet.',
    '',
    '## Highlights',
    '### Added',
    '- Overlay: Previous beta entry.',
    '',
    '## Installation',
    '',
    'See the README and docs/installation guide for full setup steps.',
    '',
  ].join('\n');

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.mkdirSync(path.join(projectRoot, 'release'), { recursive: true });
  fs.writeFileSync(
    path.join(projectRoot, 'package.json'),
    JSON.stringify({ name: 'subminer', version: '0.12.0-beta.3' }, null, 2),
    'utf8',
  );
  fs.writeFileSync(path.join(projectRoot, 'release', 'prerelease-notes.md'), existingNotes, 'utf8');
  fs.writeFileSync(
    path.join(projectRoot, 'changes', '001.md'),
    ['type: added', 'area: overlay', '', '- Added overlay coverage.'].join('\n'),
    'utf8',
  );

  try {
    const stub = defaultStubClaude();
    writePrereleaseNotesForVersion({
      cwd: projectRoot,
      version: '0.12.0-beta.3',
      deps: {
        runClaude: stub.runClaude,
        listPrereleaseTags: () => [],
        resolveFragmentDelta: () => [],
      },
    });

    assert.equal(stub.calls.length, 1);
    const prompt = stub.calls[0]!.input;
    assert.match(prompt, /EXISTING PRERELEASE NOTES/);
    assert.match(prompt, /Overlay: Previous beta entry\./);
    assert.doesNotMatch(prompt, /Stale beta-to-beta delta bullet\./);
    assert.doesNotMatch(prompt, /## Changes since /);
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('verifyPrereleaseNotesMatchVersion accepts matching notes and rejects stale or legacy markers', async () => {
  const { verifyPrereleaseNotesMatchVersion } = await loadModule();
  const workspace = createWorkspace('verify-prerelease-notes');
  const projectRoot = path.join(workspace, 'SubMiner');
  const notesPath = path.join(projectRoot, 'release', 'prerelease-notes.md');

  fs.mkdirSync(path.join(projectRoot, 'release'), { recursive: true });
  fs.writeFileSync(
    path.join(projectRoot, 'package.json'),
    JSON.stringify({ name: 'subminer', version: '0.12.0-beta.2' }, null, 2),
    'utf8',
  );

  try {
    assert.throws(
      () => verifyPrereleaseNotesMatchVersion({ cwd: projectRoot, version: '0.12.0-beta.2' }),
      /Missing .*prerelease-notes\.md/,
    );

    fs.writeFileSync(
      notesPath,
      '<!-- prerelease-version: 0.12.0-beta.2; since: v0.12.0-beta.1 -->\n\n## Highlights\n',
      'utf8',
    );
    verifyPrereleaseNotesMatchVersion({ cwd: projectRoot, version: '0.12.0-beta.2' });
    verifyPrereleaseNotesMatchVersion({ cwd: projectRoot, version: 'v0.12.0-beta.2' });

    fs.writeFileSync(
      notesPath,
      '<!-- prerelease-version: 0.12.0-beta.1 -->\n\n## Highlights\n',
      'utf8',
    );
    assert.throws(
      () => verifyPrereleaseNotesMatchVersion({ cwd: projectRoot, version: '0.12.0-beta.2' }),
      /generated for 0\.12\.0-beta\.1 but this release is 0\.12\.0-beta\.2/,
    );

    fs.writeFileSync(
      notesPath,
      '<!-- prerelease-base-version: 0.12.0 -->\n\n## Highlights\n',
      'utf8',
    );
    assert.throws(
      () => verifyPrereleaseNotesMatchVersion({ cwd: projectRoot, version: '0.12.0-beta.2' }),
      /missing or legacy prerelease-version marker/,
    );
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

test('default git tag listing and fragment delta resolution work against a real repository', async () => {
  const { writePrereleaseNotesForVersion } = await loadModule();
  const workspace = createWorkspace('prerelease-git-defaults');
  const projectRoot = path.join(workspace, 'SubMiner');
  const git = (...args: string[]): void => {
    execFileSync('git', args, { cwd: projectRoot, stdio: 'ignore' });
  };

  fs.mkdirSync(path.join(projectRoot, 'changes'), { recursive: true });
  fs.writeFileSync(
    path.join(projectRoot, 'package.json'),
    JSON.stringify({ name: 'subminer', version: '0.11.3-beta.1' }, null, 2),
    'utf8',
  );
  fs.writeFileSync(
    path.join(projectRoot, 'changes', 'kept.md'),
    ['type: added', 'area: overlay', '', '- Kept change.'].join('\n'),
    'utf8',
  );
  fs.writeFileSync(
    path.join(projectRoot, 'changes', 'edited.md'),
    ['type: fixed', 'area: launcher', '', '- Original launcher fix.'].join('\n'),
    'utf8',
  );
  fs.writeFileSync(
    path.join(projectRoot, 'changes', 'removed.md'),
    ['type: added', 'area: stats', '', '- Reverted stats change.'].join('\n'),
    'utf8',
  );

  try {
    git('init', '--quiet');
    git('-c', 'user.email=test@example.com', '-c', 'user.name=Test', 'add', '.');
    git('-c', 'user.email=test@example.com', '-c', 'user.name=Test', 'commit', '-m', 'beta.1');
    git('tag', 'v0.11.3-beta.1');

    fs.writeFileSync(
      path.join(projectRoot, 'changes', 'edited.md'),
      ['type: fixed', 'area: launcher', '', '- Broader launcher fix.'].join('\n'),
      'utf8',
    );
    fs.rmSync(path.join(projectRoot, 'changes', 'removed.md'));
    fs.writeFileSync(
      path.join(projectRoot, 'changes', 'new.md'),
      ['type: added', 'area: anki', '', '- New anki change.'].join('\n'),
      'utf8',
    );
    fs.writeFileSync(
      path.join(projectRoot, 'package.json'),
      JSON.stringify({ name: 'subminer', version: '0.11.3-beta.2' }, null, 2),
      'utf8',
    );

    const stub = recordingRunClaude((input) =>
      input.includes('PREVIOUS_TAG:') ? '- Delta bullet.' : defaultPolishedBody(input),
    );
    writePrereleaseNotesForVersion({
      cwd: projectRoot,
      version: '0.11.3-beta.2',
      deps: { runClaude: stub.runClaude },
    });

    assert.equal(stub.calls.length, 2);
    const deltaPrompt = stub.calls[0]!.input;
    assert.match(deltaPrompt, /PREVIOUS_TAG: v0\.11\.3-beta\.1/);
    assert.match(deltaPrompt, /ADDED FRAGMENT changes\/new\.md/);
    assert.match(deltaPrompt, /- New anki change\./);
    assert.match(deltaPrompt, /MODIFIED FRAGMENT changes\/edited\.md/);
    assert.match(deltaPrompt, /- Original launcher fix\./);
    assert.match(deltaPrompt, /- Broader launcher fix\./);
    assert.match(deltaPrompt, /DELETED FRAGMENT changes\/removed\.md/);
    assert.match(deltaPrompt, /- Reverted stats change\./);
    assert.doesNotMatch(deltaPrompt, /kept\.md/);
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
});

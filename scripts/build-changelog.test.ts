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

interface Project {
  /** Workspace parent dir; also handy as an empty directory. */
  workspace: string;
  projectRoot: string;
  read(relativePath: string): string;
  exists(relativePath: string): boolean;
  write(relativePath: string, content: string): void;
  remove(relativePath: string): void;
  /** Absolute path of `relativePath` inside the project. */
  at(relativePath: string): string;
}

const EMPTY_CHANGELOG = '# Changelog\n';

function packageJson(version: string): string {
  return JSON.stringify({ name: 'subminer', version }, null, 2);
}

/** Builds a fragment file body; `extraMeta` lines (e.g. `breaking: true`) follow `area`. */
function fragment(type: string, area: string, body: string, extraMeta: string[] = []): string {
  return [`type: ${type}`, `area: ${area}`, ...extraMeta, '', body].join('\n');
}

/**
 * Creates a temp project containing `files` (relative path -> content), always with an
 * empty `changes/` dir, runs `run`, then removes the workspace.
 */
async function withProject<T>(
  name: string,
  files: Record<string, string>,
  run: (project: Project) => T | Promise<T>,
): Promise<T> {
  const workspace = createWorkspace(name);
  const projectRoot = path.join(workspace, 'SubMiner');
  const at = (relativePath: string) => path.join(projectRoot, relativePath);
  const write = (relativePath: string, content: string) => {
    fs.mkdirSync(path.dirname(at(relativePath)), { recursive: true });
    fs.writeFileSync(at(relativePath), content, 'utf8');
  };
  fs.mkdirSync(at('changes'), { recursive: true });
  for (const [relativePath, content] of Object.entries(files)) write(relativePath, content);

  try {
    return await run({
      workspace,
      projectRoot,
      at,
      write,
      read: (relativePath) => fs.readFileSync(at(relativePath), 'utf8'),
      exists: (relativePath) => fs.existsSync(at(relativePath)),
      remove: (relativePath) => fs.rmSync(at(relativePath)),
    });
  } finally {
    fs.rmSync(workspace, { recursive: true, force: true });
  }
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

const PRERELEASE_BANNER =
  '> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.';

const INSTALLATION_SECTION = [
  '## Installation',
  '',
  'See the README and docs/installation guide for full setup steps.',
  '',
];

/** Fixture of a previous prerelease-notes.md: banner, optional marker, then `bodyLines`. */
function existingPrereleaseNotes(markerLine: string | null, bodyLines: string[]): string {
  return [PRERELEASE_BANNER, '', ...(markerLine ? [markerLine, ''] : []), ...bodyLines].join('\n');
}

test('writeChangelogArtifacts ignores README, groups fragments by type, writes release notes, and deletes only fragment files', async () => {
  const { writeChangelogArtifacts } = await loadModule();
  const existingChangelog = [
    '# Changelog',
    '',
    '## v0.4.0 (2026-03-01)',
    '- Existing fix',
    '',
  ].join('\n');

  await withProject(
    'write-artifacts',
    {
      'CHANGELOG.md': existingChangelog,
      'changes/README.md': '# Changelog Fragments\n\nIgnored helper text.\n',
      'changes/001.md': fragment('added', 'overlay', '- Added release fragments.'),
      'changes/002.md': fragment('fixed', 'release', 'Fixed release notes generation.'),
    },
    (project) => {
      const stub = defaultStubClaude();
      const result = writeChangelogArtifacts({
        cwd: project.projectRoot,
        version: '0.4.1',
        date: '2026-03-07',
        deps: { runClaude: stub.runClaude },
      });

      assert.deepEqual(result.outputPaths, [project.at('CHANGELOG.md')]);
      assert.deepEqual(result.deletedFragmentPaths, [
        project.at('changes/001.md'),
        project.at('changes/002.md'),
      ]);
      assert.equal(project.exists('changes/001.md'), false);
      assert.equal(project.exists('changes/002.md'), false);
      assert.equal(project.exists('changes/README.md'), true);

      assert.deepEqual(
        stub.calls.map((call) => modeFromPrompt(call.input)),
        ['changelog', 'release-notes'],
        'expected one Claude call per output (changelog + release notes)',
      );

      const changelog = project.read('CHANGELOG.md');
      assert.match(changelog, /^# Changelog\n\n## v0\.4\.1 \(2026-03-07\)\n\n/);
      assert.match(changelog, /### Added\n- Polished: added entry\./);
      assert.match(changelog, /### Fixed\n- Polished: fixed entry\./);
      assert.match(changelog, /## v0\.4\.0 \(2026-03-01\)\n- Existing fix\n$/);

      const releaseNotes = project.read('release/release-notes.md');
      assert.match(releaseNotes, /## Highlights\n### Added\n- Polished: added entry\./);
      assert.match(releaseNotes, /### Fixed\n- Polished: fixed entry\./);
      assert.match(releaseNotes, /## Installation\n\nSee the README and docs\/installation guide/);
      assert.match(releaseNotes, /- Windows: `SubMiner-\*\.exe` and `SubMiner-\*-win\.zip`/);
    },
  );
});

test('writeChangelogArtifacts skips changelog prepend when release section already exists', async () => {
  const { writeChangelogArtifacts } = await loadModule();
  const existingChangelog = [
    '# Changelog',
    '',
    '## v0.4.1 (2026-03-07)',
    '### Added',
    '- Existing release bullet.',
    '',
  ].join('\n');

  await withProject(
    'write-artifacts-existing-version',
    {
      'CHANGELOG.md': existingChangelog,
      'changes/001.md': fragment('added', 'overlay', '- Stale release fragment.'),
    },
    (project) => {
      const result = writeChangelogArtifacts({
        cwd: project.projectRoot,
        version: '0.4.1',
        date: '2026-03-08',
      });

      assert.deepEqual(result.deletedFragmentPaths, [project.at('changes/001.md')]);
      assert.equal(project.exists('changes/001.md'), false);
      assert.equal(project.read('CHANGELOG.md'), existingChangelog);
      assert.match(
        project.read('release/release-notes.md'),
        /## Highlights\n### Added\n- Existing release bullet\./,
      );
    },
  );
});

test('writeStableReleaseArtifacts reuses the requested version and date for changelog, release notes, and docs-site output', async () => {
  const { writeStableReleaseArtifacts } = await loadModule();

  await withProject(
    'write-stable-release-artifacts',
    {
      'docs-site/.keep': '',
      'package.json': packageJson('0.4.1'),
      'CHANGELOG.md': EMPTY_CHANGELOG,
      'changes/001.md': fragment('fixed', 'release', '- Reused explicit stable release date.'),
    },
    (project) => {
      const stub = defaultStubClaude();
      const result = writeStableReleaseArtifacts({
        cwd: project.projectRoot,
        version: '0.4.1',
        date: '2026-03-07',
        deps: { runClaude: stub.runClaude },
      });

      assert.deepEqual(result.outputPaths, [project.at('CHANGELOG.md')]);
      assert.equal(result.releaseNotesPath, project.at('release/release-notes.md'));
      assert.equal(result.docsChangelogPath, project.at('docs-site/changelog.md'));

      assert.match(project.read('CHANGELOG.md'), /## v0\.4\.1 \(2026-03-07\)/);
      assert.match(project.read('docs-site/changelog.md'), /## v0\.4\.1 \(2026-03-07\)/);
    },
  );
});

test('verifyChangelogReadyForRelease ignores README but rejects pending fragments and missing version sections', async () => {
  const { verifyChangelogReadyForRelease } = await loadModule();

  await withProject(
    'verify-release',
    {
      'CHANGELOG.md': EMPTY_CHANGELOG,
      'changes/README.md': '# Changelog Fragments\n',
      'changes/001.md': '- Pending fragment.\n',
    },
    (project) => {
      assert.throws(
        () => verifyChangelogReadyForRelease({ cwd: project.projectRoot, version: '0.4.1' }),
        /Pending changelog fragments/,
      );

      project.remove('changes/001.md');

      assert.throws(
        () => verifyChangelogReadyForRelease({ cwd: project.projectRoot, version: '0.4.1' }),
        /Missing CHANGELOG section for v0\.4\.1/,
      );
    },
  );
});

test('verifyChangelogReadyForRelease rejects explicit release versions that do not match package.json', async () => {
  const { verifyChangelogReadyForRelease } = await loadModule();

  await withProject(
    'verify-release-version-match',
    {
      'package.json': packageJson('0.4.0'),
      'CHANGELOG.md': '# Changelog\n\n## v0.4.1 (2026-03-09)\n- Ready.\n',
    },
    (project) => {
      assert.throws(
        () => verifyChangelogReadyForRelease({ cwd: project.projectRoot, version: '0.4.1' }),
        /package\.json version \(0\.4\.0\) does not match requested release version \(0\.4\.1\)/,
      );
    },
  );
});

test('writeChangelogArtifacts renders breaking changes section above type sections', async () => {
  const { writeChangelogArtifacts } = await loadModule();

  await withProject(
    'breaking-changes',
    {
      'CHANGELOG.md': EMPTY_CHANGELOG,
      'changes/001.md': fragment('changed', 'config', '- Renamed `foo` to `bar`.', [
        'breaking: true',
      ]),
      'changes/002.md': fragment('fixed', 'overlay', '- Fixed subtitle rendering.'),
    },
    (project) => {
      const stub = defaultStubClaude();
      writeChangelogArtifacts({
        cwd: project.projectRoot,
        version: '0.5.0',
        date: '2026-04-06',
        deps: { runClaude: stub.runClaude },
      });

      const changelog = project.read('CHANGELOG.md');
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
    },
  );
});

test('verifyChangelogFragments rejects invalid metadata', async () => {
  const { verifyChangelogFragments } = await loadModule();

  await withProject(
    'lint-invalid',
    { 'changes/001.md': fragment('nope', 'overlay', '- Invalid type.') },
    (project) => {
      assert.throws(
        () => verifyChangelogFragments({ cwd: project.projectRoot }),
        /must declare type as one of/,
      );
    },
  );
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
  const existingChangelog = '# Changelog\n\n## v0.11.2 (2026-04-07)\n- Stable release.\n';
  const existingDocsChangelog = '# Changelog\n\n## v0.11.2 (2026-04-07)\n- Stable docs release.\n';

  await withProject(
    'prerelease-beta-notes',
    {
      'package.json': packageJson('0.11.3-beta.1'),
      'CHANGELOG.md': existingChangelog,
      'docs-site/changelog.md': existingDocsChangelog,
      'changes/001.md': fragment('added', 'overlay', '- Added prerelease coverage.'),
      'changes/002.md': fragment('fixed', 'launcher', '- Fixed prerelease packaging checks.'),
    },
    (project) => {
      const stub = defaultStubClaude();
      const outputPath = writePrereleaseNotesForVersion({
        cwd: project.projectRoot,
        version: '0.11.3-beta.1',
        deps: { runClaude: stub.runClaude, listPrereleaseTags: () => [] },
      });

      assert.equal(outputPath, project.at('release/prerelease-notes.md'));
      assert.equal(
        project.read('CHANGELOG.md'),
        existingChangelog,
        'stable CHANGELOG.md should remain unchanged',
      );
      assert.equal(
        project.read('docs-site/changelog.md'),
        existingDocsChangelog,
        'docs-site changelog should remain unchanged',
      );
      assert.equal(project.exists('changes/001.md'), true);
      assert.equal(project.exists('changes/002.md'), true);

      assert.equal(stub.calls.length, 1, 'prerelease should issue exactly one Claude call');
      assert.equal(modeFromPrompt(stub.calls[0]!.input), 'release-notes');

      const prereleaseNotes = fs.readFileSync(outputPath, 'utf8');
      assert.match(prereleaseNotes, /^> This is a prerelease build for testing\./m);
      assert.match(prereleaseNotes, /<!-- prerelease-version: 0\.11\.3-beta\.1 -->/);
      assert.doesNotMatch(prereleaseNotes, /## Changes since /);
      assert.match(prereleaseNotes, /## Highlights\n### Added\n- Polished: added entry\./);
      assert.match(prereleaseNotes, /### Fixed\n- Polished: fixed entry\./);
      assert.match(
        prereleaseNotes,
        /## Installation\n\nSee the README and docs\/installation guide/,
      );
      assert.match(prereleaseNotes, /Windows `subminer\.cmd` launcher/);
      assert.match(
        prereleaseNotes,
        /Both launcher downloads use Bun included with the SubMiner app/,
      );
      assert.match(prereleaseNotes, /Bun corresponding source: `bun-v1\.3\.5-source\.tar\.gz`/);
      assert.match(prereleaseNotes, /statically links JavaScriptCore \(LGPL 2\.0\)/);
    },
  );
});

test('writePrereleaseNotesForVersion reuses existing prerelease notes when adding new fragments', async () => {
  const { writePrereleaseNotesForVersion } = await loadModule();
  const existingNotes = existingPrereleaseNotes('<!-- prerelease-base-version: 0.11.3 -->', [
    '## Highlights',
    '### Added',
    '- Overlay: Previous beta entry.',
    '',
    ...INSTALLATION_SECTION,
    '## Assets',
    '',
    '- Linux: `SubMiner.AppImage`',
    '',
  ]);

  await withProject(
    'prerelease-reuse-existing-notes',
    {
      'package.json': packageJson('0.11.3-beta.2'),
      'release/prerelease-notes.md': existingNotes,
      'changes/001.md': fragment('fixed', 'launcher', '- Fixed launcher prerelease packaging.'),
    },
    (project) => {
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
        cwd: project.projectRoot,
        version: '0.11.3-beta.2',
        deps: { runClaude: stub.runClaude, listPrereleaseTags: () => [] },
      });

      assert.equal(stub.calls.length, 1, 'prerelease should issue exactly one Claude call');
      assert.match(stub.calls[0]!.input, /EXISTING PRERELEASE NOTES/);

      const prereleaseNotes = fs.readFileSync(outputPath, 'utf8');
      assert.match(prereleaseNotes, /- Overlay: Previous beta entry\./);
      assert.match(prereleaseNotes, /- Launcher: Added only the latest fix\./);
    },
  );
});

test('writePrereleaseNotesForVersion ignores unmarked prerelease notes from an older release line', async () => {
  const { writePrereleaseNotesForVersion } = await loadModule();
  const existingNotes = existingPrereleaseNotes(null, [
    '## Highlights',
    '### Added',
    '- Settings Window: Previous release line entry.',
    '',
    ...INSTALLATION_SECTION,
  ]);

  await withProject(
    'prerelease-ignore-unmarked-old-notes',
    {
      'package.json': packageJson('0.17.0-beta.1'),
      'release/prerelease-notes.md': existingNotes,
      'changes/001.md': fragment(
        'changed',
        'overlay',
        '- Replaced subtitle delay actions with native mpv keybindings.',
      ),
    },
    (project) => {
      const stub = defaultStubClaude();
      const outputPath = writePrereleaseNotesForVersion({
        cwd: project.projectRoot,
        version: '0.17.0-beta.1',
        deps: { runClaude: stub.runClaude, listPrereleaseTags: () => [] },
      });

      assert.equal(stub.calls.length, 1, 'prerelease should issue exactly one Claude call');
      assert.doesNotMatch(stub.calls[0]!.input, /EXISTING PRERELEASE NOTES/);
      assert.doesNotMatch(stub.calls[0]!.input, /Settings Window: Previous release line entry/);

      const prereleaseNotes = fs.readFileSync(outputPath, 'utf8');
      assert.match(prereleaseNotes, /### Changed\n- Polished: changed entry\./);
    },
  );
});

test('writePrereleaseNotesForVersion supports rc prereleases', async () => {
  const { writePrereleaseNotesForVersion } = await loadModule();

  await withProject(
    'prerelease-rc-notes',
    {
      'package.json': packageJson('0.11.3-rc.1'),
      'changes/001.md': fragment('changed', 'release', '- Prepared release candidate notes.'),
    },
    (project) => {
      const stub = defaultStubClaude();
      const outputPath = writePrereleaseNotesForVersion({
        cwd: project.projectRoot,
        version: '0.11.3-rc.1',
        deps: { runClaude: stub.runClaude, listPrereleaseTags: () => [] },
      });

      assert.match(
        fs.readFileSync(outputPath, 'utf8'),
        /## Highlights\n### Changed\n- Polished: changed entry\./,
      );
    },
  );
});

const prereleaseRejectionCases = [
  {
    name: 'rejects unsupported prerelease identifiers',
    packageVersion: '0.11.3-alpha.1',
    requestedVersion: '0.11.3-alpha.1',
    fragments: {
      'changes/001.md': fragment('added', 'overlay', '- Unsupported alpha prerelease.'),
    },
    error: /Unsupported prerelease version \(0\.11\.3-alpha\.1\)/,
  },
  {
    name: 'rejects mismatched package versions',
    packageVersion: '0.11.3-beta.1',
    requestedVersion: '0.11.3-beta.2',
    fragments: { 'changes/001.md': fragment('added', 'overlay', '- Mismatched prerelease.') },
    error:
      /package\.json version \(0\.11\.3-beta\.1\) does not match requested release version \(0\.11\.3-beta\.2\)/,
  },
  {
    name: 'rejects empty prerelease note generation when no fragments exist',
    packageVersion: '0.11.3-beta.1',
    requestedVersion: '0.11.3-beta.1',
    fragments: {},
    error: /No changelog fragments found in changes\//,
  },
];

for (const c of prereleaseRejectionCases) {
  test(`writePrereleaseNotesForVersion ${c.name}`, async () => {
    const { writePrereleaseNotesForVersion } = await loadModule();

    await withProject(
      'prerelease-reject',
      { 'package.json': packageJson(c.packageVersion), ...c.fragments },
      (project) => {
        assert.throws(
          () =>
            writePrereleaseNotesForVersion({
              cwd: project.projectRoot,
              version: c.requestedVersion,
            }),
          c.error,
        );
      },
    );
  });
}

/** Project with one valid fragment and a fixed stable-release call, for Claude failure modes. */
async function expectChangelogFailure(
  name: string,
  files: Record<string, string>,
  runClaude: (input: string, args: string[]) => string,
  error: RegExp,
): Promise<void> {
  const { writeChangelogArtifacts } = await loadModule();

  await withProject(name, { 'CHANGELOG.md': EMPTY_CHANGELOG, ...files }, (project) => {
    assert.throws(
      () =>
        writeChangelogArtifacts({
          cwd: project.projectRoot,
          version: '0.5.0',
          date: '2026-04-06',
          deps: { runClaude },
        }),
      error,
    );
  });
}

const aChange = { 'changes/001.md': fragment('added', 'overlay', '- A change.') };

test('writeChangelogArtifacts surfaces a clear error when claude is missing from PATH', async () => {
  // The production defaultRunClaude wrapper translates ENOENT into this friendly
  // message; we simulate that contract here so the test exercises the propagation
  // path through polishFragmentsWithClaude rather than re-implementing the
  // execFileSync mock.
  const enoent = (): string => {
    throw new Error(
      "claude CLI not found on PATH. Install Claude Code (https://claude.com/claude-code) and ensure 'claude' is on your PATH before running changelog:build.",
    );
  };

  await expectChangelogFailure('claude-missing', aChange, enoent, /claude CLI not found on PATH/);
});

test('writeChangelogArtifacts rejects empty claude output', async () => {
  await expectChangelogFailure(
    'claude-empty',
    aChange,
    () => '   \n  ',
    /claude returned empty output/,
  );
});

test('writeChangelogArtifacts rejects claude output missing required section headers', async () => {
  await expectChangelogFailure(
    'claude-no-headers',
    aChange,
    () => 'Sure, here is your changelog: it is great.',
    /missing the expected section heading/,
  );
});

test('writeChangelogArtifacts rejects changelog-mode output that omits the Internal <details> wrapper when internal fragments are present', async () => {
  const noDetailsResponder = (input: string): string => {
    if (modeFromPrompt(input) === 'changelog') {
      return '### Added\n- Polished: added.\n\n### Internal\n- Polished: internal (no details wrapper).';
    }
    return defaultPolishedBody(input);
  };

  await expectChangelogFailure(
    'claude-no-details',
    {
      'changes/001.md': fragment('added', 'overlay', '- A user-facing change.'),
      'changes/002.md': fragment('internal', 'release', '- An internal note.'),
    },
    noDetailsResponder,
    /<details><summary>Internal changes<\/summary> wrapper/,
  );
});

test('writeChangelogArtifacts filters internal fragments from the release-notes Claude prompt', async () => {
  const { writeChangelogArtifacts } = await loadModule();

  await withProject(
    'release-notes-internal-filter',
    {
      'CHANGELOG.md': EMPTY_CHANGELOG,
      'changes/001.md': fragment('added', 'overlay', '- A user-facing change.'),
      'changes/002.md': fragment('internal', 'release', '- An internal CI tweak.'),
    },
    (project) => {
      const stub = defaultStubClaude();
      writeChangelogArtifacts({
        cwd: project.projectRoot,
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

      const releaseNotes = project.read('release/release-notes.md');
      assert.doesNotMatch(releaseNotes, /<details>/);
      assert.doesNotMatch(releaseNotes, /### Internal/);
      assert.match(
        project.read('CHANGELOG.md'),
        /<details>[\s\S]*<summary>Internal changes<\/summary>/,
      );
    },
  );
});

test('writeChangelogArtifacts appends contributor attribution and a new-contributors section to release notes', async () => {
  const { writeChangelogArtifacts } = await loadModule();

  await withProject(
    'release-notes-contributors',
    {
      'CHANGELOG.md': EMPTY_CHANGELOG,
      'changes/001.md': fragment('added', 'overlay', '- Added a feature.'),
      'changes/002.md': fragment('fixed', 'jellyfin', '- Fixed a bug.'),
    },
    (project) => {
      const stub = defaultStubClaude();
      const resolveContributionsCalls: string[][] = [];
      writeChangelogArtifacts({
        cwd: project.projectRoot,
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
        project.at('changes/001.md'),
        project.at('changes/002.md'),
      ]);

      const releaseNotes = project.read('release/release-notes.md');
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
      const changelog = project.read('CHANGELOG.md');
      assert.doesNotMatch(changelog, /What's Changed|What’s Changed/);
      assert.doesNotMatch(changelog, /New Contributors/);
    },
  );
});

test('writeChangelogArtifacts skips contributor attribution in GitHub Actions without a token', async () => {
  const { writeChangelogArtifacts } = await loadModule();
  const envKeys = ['GITHUB_ACTIONS', 'GH_TOKEN', 'GITHUB_TOKEN', 'PATH'] as const;
  const originalEnv = Object.fromEntries(envKeys.map((key) => [key, process.env[key]]));
  const originalWarn = console.warn;
  const warnings: string[] = [];

  await withProject(
    'release-notes-actions-no-token',
    {
      'CHANGELOG.md': EMPTY_CHANGELOG,
      'changes/001.md': fragment('added', 'release', '- Added a feature.'),
    },
    (project) => {
      try {
        process.env.GITHUB_ACTIONS = 'true';
        delete process.env.GH_TOKEN;
        delete process.env.GITHUB_TOKEN;
        // An empty PATH dir means the default lookup could not find `gh` even if it tried.
        process.env.PATH = project.workspace;
        console.warn = (message?: unknown) => {
          warnings.push(String(message));
        };

        writeChangelogArtifacts({
          cwd: project.projectRoot,
          version: '0.6.0',
          date: '2026-05-06',
          deps: { runClaude: defaultStubClaude().runClaude },
        });

        assert.deepEqual(warnings, []);
        assert.doesNotMatch(project.read('release/release-notes.md'), /## What's Changed/);
      } finally {
        console.warn = originalWarn;
        for (const key of envKeys) {
          const original = originalEnv[key];
          if (original === undefined) delete process.env[key];
          else process.env[key] = original;
        }
      }
    },
  );
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

const INTERNAL_DETAILS_BLOCK = [
  '<details>',
  '<summary>Internal changes</summary>',
  '',
  '### Internal',
  '- Polished: internal release note.',
  '',
  '</details>',
  '',
];

test('writeReleaseNotesForVersion preserves committed contributor attribution before installation', async () => {
  const { writeReleaseNotesForVersion } = await loadModule();
  const existingChangelog = [
    '# Changelog',
    '',
    '## v0.8.0 (2026-04-17)',
    '### Added',
    '- Polished: released feature.',
    '',
    ...INTERNAL_DETAILS_BLOCK,
  ].join('\n');
  const committedReleaseNotes = [
    '## Highlights',
    '### Added',
    '- Old generated body.',
    '',
    ...INSTALLATION_SECTION,
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

  await withProject(
    'release-notes-preserve-attribution',
    { 'CHANGELOG.md': existingChangelog, 'release/release-notes.md': committedReleaseNotes },
    (project) => {
      const outputPath = writeReleaseNotesForVersion({
        cwd: project.projectRoot,
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
    },
  );
});

test('writeChangelogArtifacts strips <details> blocks from release notes when reusing an existing CHANGELOG section', async () => {
  const { writeChangelogArtifacts } = await loadModule();
  const existingChangelog = [
    '# Changelog',
    '',
    '## v0.4.1 (2026-03-07)',
    '### Added',
    '- Polished: previously committed.',
    '',
    ...INTERNAL_DETAILS_BLOCK,
  ].join('\n');

  await withProject(
    'reuse-existing-section',
    {
      'CHANGELOG.md': existingChangelog,
      'changes/001.md': fragment('added', 'overlay', '- Stale fragment.'),
    },
    (project) => {
      const stub = defaultStubClaude();
      writeChangelogArtifacts({
        cwd: project.projectRoot,
        version: '0.4.1',
        date: '2026-03-08',
        deps: { runClaude: stub.runClaude },
      });

      assert.equal(
        stub.calls.length,
        0,
        'no Claude calls should fire when the section already exists',
      );

      const releaseNotes = project.read('release/release-notes.md');
      assert.match(releaseNotes, /## Highlights\n### Added\n- Polished: previously committed\./);
      assert.doesNotMatch(releaseNotes, /<details>/);
      assert.doesNotMatch(releaseNotes, /### Internal/);
    },
  );
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

  await withProject(
    'prerelease-delta-section',
    {
      'package.json': packageJson('0.12.0-beta.2'),
      'changes/001.md': fragment('fixed', 'overlay', '- Fixed overlay focus and macOS helper.'),
    },
    (project) => {
      const stub = recordingRunClaude((input) =>
        input.includes('MODIFIED FRAGMENT')
          ? '- Fixed the macOS helper deployment target for older systems.'
          : '### Fixed\n- Overlay: cumulative fixed entry.',
      );
      const outputPath = writePrereleaseNotesForVersion({
        cwd: project.projectRoot,
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
      assert.match(
        prereleaseNotes,
        /- Fixed the macOS helper deployment target for older systems\./,
      );
    },
  );
});

test('writePrereleaseNotesForVersion renders a fallback delta line when no fragments changed', async () => {
  const { writePrereleaseNotesForVersion } = await loadModule();

  await withProject(
    'prerelease-empty-delta',
    {
      'package.json': packageJson('0.12.0-beta.3'),
      'changes/001.md': fragment('fixed', 'overlay', '- Fixed overlay focus.'),
    },
    (project) => {
      const stub = defaultStubClaude();
      const outputPath = writePrereleaseNotesForVersion({
        cwd: project.projectRoot,
        version: '0.12.0-beta.3',
        deps: {
          runClaude: stub.runClaude,
          listPrereleaseTags: () => ['v0.12.0-beta.1', 'v0.12.0-beta.2'],
          resolveFragmentDelta: () => [],
        },
      });

      assert.equal(stub.calls.length, 1, 'empty delta must not spend a Claude call');
      assert.match(
        fs.readFileSync(outputPath, 'utf8'),
        /## Changes since v0\.12\.0-beta\.2\n\n- No changelog fragment changes since v0\.12\.0-beta\.2; this build contains packaging or internal-only updates\./,
      );
    },
  );
});

test('writePrereleaseNotesForVersion rejects non-bullet delta output from Claude', async () => {
  const { writePrereleaseNotesForVersion } = await loadModule();

  await withProject(
    'prerelease-delta-invalid-output',
    {
      'package.json': packageJson('0.12.0-beta.2'),
      'changes/001.md': fragment('fixed', 'overlay', '- Fixed overlay focus.'),
    },
    (project) => {
      const stub = recordingRunClaude(() => 'Here are the changes:\n- One change.');
      assert.throws(
        () =>
          writePrereleaseNotesForVersion({
            cwd: project.projectRoot,
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
    },
  );
});

test('writePrereleaseNotesForVersion strips the stale delta section from the reused baseline', async () => {
  const { writePrereleaseNotesForVersion } = await loadModule();
  const existingNotes = existingPrereleaseNotes(
    '<!-- prerelease-version: 0.12.0-beta.2; since: v0.12.0-beta.1 -->',
    [
      '## Changes since v0.12.0-beta.1',
      '',
      '- Stale beta-to-beta delta bullet.',
      '',
      '## Highlights',
      '### Added',
      '- Overlay: Previous beta entry.',
      '',
      ...INSTALLATION_SECTION,
    ],
  );

  await withProject(
    'prerelease-reuse-strips-delta',
    {
      'package.json': packageJson('0.12.0-beta.3'),
      'release/prerelease-notes.md': existingNotes,
      'changes/001.md': fragment('added', 'overlay', '- Added overlay coverage.'),
    },
    (project) => {
      const stub = defaultStubClaude();
      writePrereleaseNotesForVersion({
        cwd: project.projectRoot,
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
    },
  );
});

test('verifyPrereleaseNotesMatchVersion accepts matching notes and rejects stale or legacy markers', async () => {
  const { verifyPrereleaseNotesMatchVersion } = await loadModule();

  await withProject(
    'verify-prerelease-notes',
    { 'package.json': packageJson('0.12.0-beta.2') },
    (project) => {
      const verify = (version = '0.12.0-beta.2') =>
        verifyPrereleaseNotesMatchVersion({ cwd: project.projectRoot, version });
      const notesPath = 'release/prerelease-notes.md';

      assert.throws(() => verify(), /Missing .*prerelease-notes\.md/);

      project.write(
        notesPath,
        '<!-- prerelease-version: 0.12.0-beta.2; since: v0.12.0-beta.1 -->\n\n## Highlights\n',
      );
      verify();
      verify('v0.12.0-beta.2');

      project.write(notesPath, '<!-- prerelease-version: 0.12.0-beta.1 -->\n\n## Highlights\n');
      assert.throws(
        () => verify(),
        /generated for 0\.12\.0-beta\.1 but this release is 0\.12\.0-beta\.2/,
      );

      project.write(notesPath, '<!-- prerelease-base-version: 0.12.0 -->\n\n## Highlights\n');
      assert.throws(() => verify(), /missing or legacy prerelease-version marker/);
    },
  );
});

test('default git tag listing and fragment delta resolution work against a real repository', async () => {
  const { writePrereleaseNotesForVersion } = await loadModule();

  await withProject(
    'prerelease-git-defaults',
    {
      'package.json': packageJson('0.11.3-beta.1'),
      'changes/kept.md': fragment('added', 'overlay', '- Kept change.'),
      'changes/edited.md': fragment('fixed', 'launcher', '- Original launcher fix.'),
      'changes/removed.md': fragment('added', 'stats', '- Reverted stats change.'),
    },
    (project) => {
      const git = (...args: string[]): void => {
        execFileSync('git', args, { cwd: project.projectRoot, stdio: 'ignore' });
      };

      git('init', '--quiet');
      git('-c', 'user.email=test@example.com', '-c', 'user.name=Test', 'add', '.');
      git('-c', 'user.email=test@example.com', '-c', 'user.name=Test', 'commit', '-m', 'beta.1');
      git('tag', 'v0.11.3-beta.1');

      project.write('changes/edited.md', fragment('fixed', 'launcher', '- Broader launcher fix.'));
      project.remove('changes/removed.md');
      project.write('changes/new.md', fragment('added', 'anki', '- New anki change.'));
      project.write('package.json', packageJson('0.11.3-beta.2'));

      const stub = recordingRunClaude((input) =>
        input.includes('PREVIOUS_TAG:') ? '- Delta bullet.' : defaultPolishedBody(input),
      );
      writePrereleaseNotesForVersion({
        cwd: project.projectRoot,
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
    },
  );
});

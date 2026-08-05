import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';

import { parseChangelog, resolveChangelogGroupKey } from './changelog-parse';

const SAMPLE = `# Changelog

## v0.19.2 (2026-08-04)

### Changed
- Subsync: picks both tracks now.

### Fixed
- Overlay: shows the plain line immediately.

<details>
<summary>Internal changes</summary>

### Internal
- Patched \`undici\`.

</details>

## v0.19.1 (2026-08-01)

### Added
- **Word Card Type:**
  - Adds a setting.
  - Flags clear each other.

## v0.18.0 (2026-07-01)

### Fixed
- Something older.
`;

test('changelog parser reads versions, dates, and sections in file order', () => {
  const entries = parseChangelog(SAMPLE);

  assert.deepEqual(
    entries.map((entry) => `${entry.version}@${entry.date}`),
    ['0.19.2@2026-08-04', '0.19.1@2026-08-01', '0.18.0@2026-07-01'],
  );
  assert.deepEqual(
    entries[0]?.sections.map((section) => section.heading),
    ['Changed', 'Fixed', 'Internal'],
  );
  assert.deepEqual(entries[0]?.sections[1]?.items, [
    { text: 'Overlay: shows the plain line immediately.', children: [] },
  ]);
});

test('changelog parser flags sections inside the details block as internal', () => {
  const entries = parseChangelog(SAMPLE);
  const sections = entries[0]?.sections ?? [];

  assert.deepEqual(
    sections.map((section) => section.internal),
    [false, false, true],
  );
  assert.deepEqual(sections[2]?.items, [{ text: 'Patched `undici`.', children: [] }]);
});

test('changelog parser groups entries by major.minor', () => {
  const entries = parseChangelog(SAMPLE);

  assert.deepEqual(
    entries.map((entry) => entry.groupKey),
    ['0.19', '0.19', '0.18'],
  );
  assert.equal(resolveChangelogGroupKey('1.2.3'), '1.2');
});

test('changelog parser keeps bullets that precede any section heading', () => {
  const entries = parseChangelog('## v0.1.0 (2025-01-01)\n\n- Initial release.\n');

  assert.deepEqual(entries[0]?.sections, [
    {
      heading: 'Changes',
      items: [{ text: 'Initial release.', children: [] }],
      internal: false,
    },
  ]);
});

test('changelog parser drops empty sections and tolerates missing dates', () => {
  const entries = parseChangelog('## v0.2.0\n\n### Added\n\n### Fixed\n- One fix.\n');

  assert.equal(entries[0]?.date, '');
  assert.deepEqual(
    entries[0]?.sections.map((section) => section.heading),
    ['Fixed'],
  );
});

test('changelog parser keeps indented sub-bullets nested under their lead bullet', () => {
  const entries = parseChangelog(SAMPLE);
  const added = entries[1]?.sections.find((section) => section.heading === 'Added');

  assert.deepEqual(added?.items, [
    {
      text: '**Word Card Type:**',
      children: [
        { text: 'Adds a setting.', children: [] },
        { text: 'Flags clear each other.', children: [] },
      ],
    },
  ]);
});

test('changelog parser nests three bullet levels and rejoins wrapped lines', () => {
  const entries = parseChangelog(
    [
      '## v0.9.0 (2025-05-05)',
      '',
      '### Added',
      '- Top level',
      '  - Second level',
      '    - Third level',
      '      continued on the next line',
      '  - Back to second level',
      '- Another top level',
      '',
    ].join('\n'),
  );

  assert.deepEqual(entries[0]?.sections[0]?.items, [
    {
      text: 'Top level',
      children: [
        {
          text: 'Second level',
          children: [{ text: 'Third level continued on the next line', children: [] }],
        },
        { text: 'Back to second level', children: [] },
      ],
    },
    { text: 'Another top level', children: [] },
  ]);
});

test('changelog parser reads prerelease and build metadata version headings', () => {
  const entries = parseChangelog(
    [
      '## v0.16.0 (2026-06-01)',
      '',
      '### Added',
      '- New in 0.16.',
      '',
      '## v0.15.0-rc.1+build.2 (2026-05-29)',
      '',
      '### Added',
      '- Release candidate note.',
      '',
    ].join('\n'),
  );

  // An unrecognized heading does not just vanish: its notes fold into the
  // previous release, so the version list has to stay exact.
  assert.deepEqual(
    entries.map((entry) => entry.version),
    ['0.16.0', '0.15.0-rc.1+build.2'],
  );
  assert.equal(entries[1]?.date, '2026-05-29');
  assert.equal(entries[1]?.groupKey, '0.15');
  assert.equal(entries[0]?.sections.length, 1);
});

test('changelog parser handles the repo CHANGELOG.md', () => {
  const markdown = fs.readFileSync(path.join(process.cwd(), 'CHANGELOG.md'), 'utf8');
  const entries = parseChangelog(markdown);

  assert.ok(entries.length > 3);
  for (const entry of entries) {
    assert.match(entry.version, /^\d+\.\d+\.\d+/);
    assert.ok(entry.sections.length > 0, `expected sections for v${entry.version}`);
    for (const section of entry.sections) {
      for (const item of section.items) {
        assert.ok(item.text.length > 0, `empty bullet in v${entry.version}`);
      }
    }
  }

  // Older entries group notes under a bold lead bullet; nesting must survive.
  const breaking = entries
    .find((entry) => entry.version === '0.15.0')
    ?.sections.find((section) => section.heading === 'Breaking Changes');
  assert.deepEqual(
    breaking?.items.map((item) => `${item.text}:${item.children.length}`),
    ['**Subsync:**:2', '**N+1 Highlighting:**:2'],
  );
});

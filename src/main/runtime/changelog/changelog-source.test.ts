import assert from 'node:assert/strict';
import test from 'node:test';

import { buildRawChangelogUrl, createChangelogSource } from './changelog-source';
import { buildChangelogSnapshot } from './changelog-snapshot';

const REMOTE = `# Changelog

## v0.20.0 (2026-09-01)

### Added
- Remote only entry.

## v0.19.2 (2026-08-04)

### Fixed
- Installed entry.
`;

const BUNDLED = `# Changelog

## v0.19.2 (2026-08-04)

### Fixed
- Installed entry.
`;

function createDeps(overrides: Partial<Parameters<typeof createChangelogSource>[0]> = {}) {
  return {
    fetchLatestReleaseTag: async () => 'v0.20.0',
    fetchText: async () => REMOTE,
    readBundledChangelog: () => BUNDLED,
    getInstalledVersion: () => '0.19.2',
    now: () => 1_000,
    logWarn: () => {},
    ...overrides,
  };
}

test('changelog source reads the changelog at the latest release tag', async () => {
  const urls: string[] = [];
  const source = createChangelogSource(
    createDeps({
      fetchText: async (url: string) => {
        urls.push(url);
        return REMOTE;
      },
    }),
  );

  const snapshot = await source.getSnapshot();

  assert.deepEqual(urls, [
    'https://raw.githubusercontent.com/ksyasuda/SubMiner/v0.20.0/CHANGELOG.md',
  ]);
  assert.equal(snapshot.source, 'remote');
  assert.equal(snapshot.releaseTag, 'v0.20.0');
  assert.equal(snapshot.latestVersion, '0.20.0');
  assert.equal(snapshot.installedVersion, '0.19.2');
  assert.equal(snapshot.expandedGroupKey, '0.20');
});

test('changelog source falls back to the default branch when no release tag resolves', async () => {
  const urls: string[] = [];
  const source = createChangelogSource(
    createDeps({
      fetchLatestReleaseTag: async () => null,
      fetchText: async (url: string) => {
        urls.push(url);
        return REMOTE;
      },
    }),
  );

  const snapshot = await source.getSnapshot();

  assert.deepEqual(urls, ['https://raw.githubusercontent.com/ksyasuda/SubMiner/main/CHANGELOG.md']);
  assert.equal(snapshot.source, 'remote');
  assert.equal(snapshot.releaseTag, undefined);
});

test('changelog source falls back to the bundled changelog when the download fails', async () => {
  const warnings: string[] = [];
  const source = createChangelogSource(
    createDeps({
      fetchText: async () => {
        throw new Error('offline');
      },
      logWarn: (message: string) => warnings.push(message),
    }),
  );

  const snapshot = await source.getSnapshot();

  assert.equal(snapshot.source, 'bundled');
  assert.match(snapshot.warning ?? '', /offline/);
  assert.equal(snapshot.latestVersion, '0.19.2');
  assert.equal(warnings.length, 1);
});

test('changelog source reports an error when no changelog can be loaded', async () => {
  const source = createChangelogSource(
    createDeps({
      fetchText: async () => {
        throw new Error('offline');
      },
      readBundledChangelog: () => null,
    }),
  );

  const snapshot = await source.getSnapshot();

  assert.deepEqual(snapshot.entries, []);
  assert.match(snapshot.error ?? '', /offline/);
  assert.equal(snapshot.installedVersion, '0.19.2');
});

test('changelog source caches remote results and refreshes on demand', async () => {
  let fetches = 0;
  let clock = 0;
  const source = createChangelogSource(
    createDeps({
      now: () => clock,
      fetchText: async () => {
        fetches += 1;
        return REMOTE;
      },
    }),
  );

  await source.getSnapshot();
  await source.getSnapshot();
  assert.equal(fetches, 1);

  await source.getSnapshot({ refresh: true });
  assert.equal(fetches, 2);

  clock = 11 * 60 * 1000;
  await source.getSnapshot();
  assert.equal(fetches, 3);
});

test('changelog source retries the network after a bundled fallback', async () => {
  let fetches = 0;
  const source = createChangelogSource(
    createDeps({
      fetchText: async () => {
        fetches += 1;
        throw new Error('offline');
      },
    }),
  );

  await source.getSnapshot();
  await source.getSnapshot();

  assert.equal(fetches, 2);
});

test('changelog source treats an empty remote changelog as a failure', async () => {
  const source = createChangelogSource(createDeps({ fetchText: async () => '   ' }));

  const snapshot = await source.getSnapshot();

  assert.equal(snapshot.source, 'bundled');
});

test('changelog source falls back when the remote body parses to no releases', () => {
  const warnings: string[] = [];
  const source = createChangelogSource(
    createDeps({
      // A 200 that is not a changelog, e.g. a redirect landing page.
      fetchText: async () => '<!doctype html><html><body>Moved</body></html>',
      logWarn: (message: string) => warnings.push(message),
    }),
  );

  return source.getSnapshot().then((snapshot) => {
    assert.equal(snapshot.source, 'bundled');
    assert.equal(snapshot.entries.length, 1);
    assert.match(snapshot.warning ?? '', /no releases/);
    assert.equal(warnings.length, 1);
  });
});

test('raw changelog urls encode the release ref', () => {
  assert.equal(
    buildRawChangelogUrl('v1.0.0', 'owner', 'repo'),
    'https://raw.githubusercontent.com/owner/repo/v1.0.0/CHANGELOG.md',
  );
});

test('snapshot expansion uses the newest version even when file order is unsorted', () => {
  const snapshot = buildChangelogSnapshot(
    '## v0.18.0 (2026-01-01)\n\n### Fixed\n- Old.\n\n## v0.19.0 (2026-02-01)\n\n### Fixed\n- New.\n',
    { installedVersion: '0.18.0', source: 'bundled' },
  );

  assert.equal(snapshot.latestVersion, '0.19.0');
  assert.equal(snapshot.expandedGroupKey, '0.19');
});

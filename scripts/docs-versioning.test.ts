import { describe, expect, test } from 'bun:test';
import {
  buildVersionManifest,
  compareStableVersionsDesc,
  versionArchiveCacheKey,
  isStableReleaseTag,
  stableTagsWithDocs,
  versionArchiveCacheName,
  versionOutputPath,
  versionPath,
} from './docs-versioning';

describe('docs versioning helpers', () => {
  test('stable tag filtering excludes beta and rc tags', () => {
    expect(isStableReleaseTag('v0.14.0')).toBe(true);
    expect(isStableReleaseTag('v0.15.0-beta.3')).toBe(false);
    expect(isStableReleaseTag('v0.15.0-rc.1')).toBe(false);
  });

  test('latest stable resolves to v0.14.0 when beta tags are present', () => {
    const tags = ['v0.13.0', 'v0.15.0-beta.3', 'v0.14.0'].sort(compareStableVersionsDesc);

    expect(tags[0]).toBe('v0.14.0');
  });

  test('tags before docs-site are skipped', () => {
    const tags = ['v0.12.0', 'v0.13.0', 'v0.14.0'];
    const hasDocsSite = (tag: string) => tag !== 'v0.12.0';

    expect(stableTagsWithDocs(tags, hasDocsSite)).toEqual(['v0.14.0', 'v0.13.0']);
  });

  test('version manifest paths are normalized', () => {
    expect(versionPath('v0.14.0')).toBe('/v/0.14.0/');
    expect(
      buildVersionManifest({
        latestStable: 'v0.14.0',
        stableVersions: ['v0.14.0'],
      }),
    ).toEqual({
      latestStable: 'v0.14.0',
      channels: [
        { label: 'Latest stable', path: '/' },
        { label: 'main', path: '/main/' },
      ],
      versions: [{ version: 'v0.14.0', path: '/v/0.14.0/' }],
    });
  });

  test('archive cache names are normalized by version and shared internals hash', () => {
    expect(versionArchiveCacheName('v0.14.0', 'abcdef1234567890')).toBe('abcdef123456-v0.14.0');
  });

  test('archive cache keys change when manifest contents change', () => {
    const firstKey = versionArchiveCacheKey({
      sharedInternalsHash: 'abcdef1234567890',
      manifestJson: '{"latestStable":"v0.14.0"}',
    });
    const secondKey = versionArchiveCacheKey({
      sharedInternalsHash: 'abcdef1234567890',
      manifestJson: '{"latestStable":"v0.15.0"}',
    });

    expect(firstKey).not.toBe(secondKey);
  });

  test('archive output paths stay relative for filesystem joins', () => {
    expect(versionOutputPath('v0.14.0')).toBe('v/0.14.0');
  });
});

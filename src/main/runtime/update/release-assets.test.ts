import test from 'node:test';
import assert from 'node:assert/strict';
import {
  compareSemverLike,
  findReleaseAsset,
  parseSha256Sums,
  selectLatestStableRelease,
} from './release-assets';

test('parseSha256Sums maps release asset basenames to hashes', () => {
  const sums = parseSha256Sums(`
1111111111111111111111111111111111111111111111111111111111111111  SubMiner.AppImage
2222222222222222222222222222222222222222222222222222222222222222 *subminer
`);

  assert.equal(
    sums.get('SubMiner.AppImage'),
    '1111111111111111111111111111111111111111111111111111111111111111',
  );
  assert.equal(
    sums.get('subminer'),
    '2222222222222222222222222222222222222222222222222222222222222222',
  );
});

test('selectLatestStableRelease ignores drafts and prereleases', () => {
  const release = selectLatestStableRelease([
    { tag_name: 'v0.16.0-beta.1', draft: false, prerelease: true, assets: [] },
    { tag_name: 'v0.15.0', draft: true, prerelease: false, assets: [] },
    { tag_name: 'v0.14.1', draft: false, prerelease: false, assets: [] },
  ]);

  assert.equal(release?.tag_name, 'v0.14.1');
});

test('selectLatestStableRelease can opt into prerelease releases', () => {
  const release = selectLatestStableRelease(
    [
      { tag_name: 'v0.16.0-beta.1', draft: false, prerelease: true, assets: [] },
      { tag_name: 'v0.15.0', draft: false, prerelease: false, assets: [] },
    ],
    'prerelease',
  );

  assert.equal(release?.tag_name, 'v0.16.0-beta.1');
});

test('compareSemverLike orders prerelease identifiers within the same base version', () => {
  assert.equal(compareSemverLike('0.15.0-beta.2', '0.15.0-beta.1') > 0, true);
  assert.equal(compareSemverLike('0.15.0-rc.1', '0.15.0-beta.2') > 0, true);
  assert.equal(compareSemverLike('0.15.0', '0.15.0-rc.1') > 0, true);
});

test('findReleaseAsset finds exact asset names only', () => {
  const release = {
    tag_name: 'v0.14.1',
    draft: false,
    prerelease: false,
    assets: [
      { name: 'subminer', browser_download_url: 'https://example.test/subminer' },
      { name: 'subminer-assets.tar.gz', browser_download_url: 'https://example.test/assets' },
    ],
  };

  assert.equal(
    findReleaseAsset(release, 'subminer')?.browser_download_url,
    'https://example.test/subminer',
  );
  assert.equal(findReleaseAsset(release, 'latest.yml'), null);
});

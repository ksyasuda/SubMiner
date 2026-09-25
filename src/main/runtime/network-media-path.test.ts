import assert from 'node:assert/strict';
import test from 'node:test';
import { createRemoteMediaPathDetector } from './network-media-path';

test('remote media detector recognizes mounted network filesystems', async () => {
  const detectRemoteMedia = createRemoteMediaPathDetector({
    platform: 'darwin',
    readMountOutput: async () =>
      [
        '/dev/disk3s5 on /System/Volumes/Data (apfs, local, journaled)',
        '//viewer@media/jellyfin on /Volumes/jellyfin (smbfs, nodev, nosuid)',
      ].join('\n'),
  });

  assert.equal(await detectRemoteMedia('/Volumes/jellyfin/movie.mkv'), true);
  assert.equal(await detectRemoteMedia('/Volumes/jellyfin-another/movie.mkv'), false);
  assert.equal(await detectRemoteMedia('/Users/viewer/movie.mkv'), false);
});

test('remote media detector recognizes Linux network mount output', async () => {
  const detectRemoteMedia = createRemoteMediaPathDetector({
    platform: 'linux',
    readMountOutput: async () =>
      '//media/jellyfin on /mnt/Jellyfin\\040Media type cifs (rw,relatime)',
  });

  assert.equal(await detectRemoteMedia('/mnt/Jellyfin Media/movie.mkv'), true);
});

test('remote media detector shares its mount lookup between concurrent callers', async () => {
  let mountReads = 0;
  const detectRemoteMedia = createRemoteMediaPathDetector({
    platform: 'darwin',
    readMountOutput: async () => {
      mountReads += 1;
      return '//viewer@media/jellyfin on /Volumes/jellyfin (smbfs, nodev, nosuid)';
    },
  });

  const results = await Promise.all(
    Array.from({ length: 6 }, () => detectRemoteMedia('/Volumes/jellyfin/movie.mkv')),
  );

  assert.deepEqual(
    results,
    Array.from({ length: 6 }, () => true),
  );
  assert.equal(mountReads, 1);
});

test('remote media detector recognizes URLs and Windows UNC paths without reading mounts', async () => {
  let mountReads = 0;
  const detectRemoteMedia = createRemoteMediaPathDetector({
    platform: 'win32',
    readMountOutput: async () => {
      mountReads += 1;
      return '';
    },
  });

  assert.equal(await detectRemoteMedia('https://media.example/movie.mkv'), true);
  assert.equal(await detectRemoteMedia('\\\\media-server\\jellyfin\\movie.mkv'), true);
  assert.equal(mountReads, 0);
});

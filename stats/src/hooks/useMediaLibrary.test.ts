import assert from 'node:assert/strict';
import test from 'node:test';
import type { MediaLibraryItem } from '../types/stats';
import { shouldRefreshMediaLibraryRows } from './useMediaLibrary';

const baseItem: MediaLibraryItem = {
  videoId: 1,
  canonicalTitle: 'watch?v=abc123',
  totalSessions: 1,
  totalActiveMs: 60_000,
  totalCards: 0,
  totalTokensSeen: 10,
  lastWatchedMs: 1_000,
  hasCoverArt: 0,
  youtubeVideoId: 'abc123',
  videoUrl: 'https://www.youtube.com/watch?v=abc123',
  videoTitle: null,
  videoThumbnailUrl: 'https://i.ytimg.com/vi/abc123/hqdefault.jpg',
  channelId: null,
  channelName: null,
  channelUrl: null,
  channelThumbnailUrl: null,
  uploaderId: null,
  uploaderUrl: null,
  description: null,
};

test('shouldRefreshMediaLibraryRows requests a follow-up fetch for incomplete youtube metadata', () => {
  assert.equal(shouldRefreshMediaLibraryRows([baseItem]), true);
});

test('shouldRefreshMediaLibraryRows skips follow-up fetch when youtube metadata is complete', () => {
  assert.equal(
    shouldRefreshMediaLibraryRows([
      {
        ...baseItem,
        videoTitle: 'Video Name',
        channelName: 'Creator Name',
        channelThumbnailUrl: 'https://yt3.googleusercontent.com/channel-avatar=s88',
      },
    ]),
    false,
  );
});

test('shouldRefreshMediaLibraryRows ignores non-youtube rows', () => {
  assert.equal(
    shouldRefreshMediaLibraryRows([
      {
        ...baseItem,
        youtubeVideoId: null,
        videoUrl: null,
      },
    ]),
    false,
  );
});

test('useMediaLibrary is a function export', async () => {
  // Verify the hook is exported as a function. The `refresh` return value
  // is exercised at the component level; reactive re-fetch via version bump
  // cannot be tested without a full React test environment (no @testing-library).
  const mod = await import('./useMediaLibrary');
  assert.equal(typeof mod.useMediaLibrary, 'function');
});

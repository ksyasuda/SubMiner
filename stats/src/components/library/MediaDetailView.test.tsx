import assert from 'node:assert/strict';
import test from 'node:test';
import { getRelatedCollectionLabel } from './MediaDetailView';

test('getRelatedCollectionLabel returns View Channel for youtube-backed media', () => {
  assert.equal(
    getRelatedCollectionLabel({
      videoId: 1,
      animeId: 1,
      canonicalTitle: 'Video',
      totalSessions: 1,
      totalActiveMs: 1,
      totalCards: 0,
      totalTokensSeen: 0,
      totalLinesSeen: 0,
      totalLookupCount: 0,
      totalLookupHits: 0,
      totalYomitanLookupCount: 0,
      channelName: 'Creator',
    }),
    'View Channel',
  );
});

test('getRelatedCollectionLabel returns View Anime for non-youtube media', () => {
  assert.equal(
    getRelatedCollectionLabel({
      videoId: 2,
      animeId: 1,
      canonicalTitle: 'Episode 5',
      totalSessions: 1,
      totalActiveMs: 1,
      totalCards: 0,
      totalTokensSeen: 0,
      totalLinesSeen: 0,
      totalLookupCount: 0,
      totalLookupHits: 0,
      totalYomitanLookupCount: 0,
      channelName: null,
    }),
    'View Anime',
  );
});

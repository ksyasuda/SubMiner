import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { createElement } from 'react';
import { getRelatedCollectionLabel, buildDeleteEpisodeHandler } from './MediaDetailView';

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

test('buildDeleteEpisodeHandler calls deleteVideo then onBack when confirm returns true', async () => {
  let deletedVideoId: number | null = null;
  let onBackCalled = false;

  const fakeApiClient = {
    deleteVideo: async (id: number) => {
      deletedVideoId = id;
    },
  };

  const fakeConfirm = (_title: string) => true;

  const handler = buildDeleteEpisodeHandler({
    videoId: 42,
    title: 'Test Episode',
    apiClient: fakeApiClient as { deleteVideo: (id: number) => Promise<void> },
    confirmFn: fakeConfirm,
    onBack: () => {
      onBackCalled = true;
    },
    setDeleteError: () => {},
  });

  await handler();
  assert.equal(deletedVideoId, 42);
  assert.equal(onBackCalled, true);
});

test('buildDeleteEpisodeHandler does nothing when confirm returns false', async () => {
  let deletedVideoId: number | null = null;
  let onBackCalled = false;

  const fakeApiClient = {
    deleteVideo: async (id: number) => {
      deletedVideoId = id;
    },
  };

  const fakeConfirm = (_title: string) => false;

  const handler = buildDeleteEpisodeHandler({
    videoId: 42,
    title: 'Test Episode',
    apiClient: fakeApiClient as { deleteVideo: (id: number) => Promise<void> },
    confirmFn: fakeConfirm,
    onBack: () => {
      onBackCalled = true;
    },
    setDeleteError: () => {},
  });

  await handler();
  assert.equal(deletedVideoId, null);
  assert.equal(onBackCalled, false);
});

test('buildDeleteEpisodeHandler sets error when deleteVideo throws', async () => {
  let capturedError: string | null = null;

  const fakeApiClient = {
    deleteVideo: async (_id: number) => {
      throw new Error('Network failure');
    },
  };

  const fakeConfirm = (_title: string) => true;

  const handler = buildDeleteEpisodeHandler({
    videoId: 42,
    title: 'Test Episode',
    apiClient: fakeApiClient as { deleteVideo: (id: number) => Promise<void> },
    confirmFn: fakeConfirm,
    onBack: () => {},
    setDeleteError: (msg) => {
      capturedError = msg;
    },
  });

  await handler();
  assert.equal(capturedError, 'Network failure');
});

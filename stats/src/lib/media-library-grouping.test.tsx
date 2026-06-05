import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import type { MediaLibraryItem } from '../types/stats';
import {
  groupMediaLibraryItems,
  resolveMediaArtworkUrl,
  resolveMediaCoverApiUrl,
  summarizeMediaLibraryGroups,
} from './media-library-grouping';
import { CoverImage } from '../components/library/CoverImage';
import { MediaCard } from '../components/library/MediaCard';

const youtubeEpisodeA: MediaLibraryItem = {
  videoId: 1,
  canonicalTitle: 'Episode 1',
  totalSessions: 2,
  totalActiveMs: 12_000,
  totalCards: 3,
  totalTokensSeen: 120,
  lastWatchedMs: 3_000,
  hasCoverArt: 1,
  youtubeVideoId: 'yt-1',
  videoUrl: 'https://www.youtube.com/watch?v=yt-1',
  videoTitle: 'Video 1',
  videoThumbnailUrl: 'https://i.ytimg.com/vi/yt-1/hqdefault.jpg',
  channelId: 'UC123',
  channelName: 'Creator Name',
  channelUrl: 'https://www.youtube.com/channel/UC123',
  channelThumbnailUrl: 'https://yt3.googleusercontent.com/channel-avatar=s88',
  uploaderId: '@creator',
  uploaderUrl: 'https://www.youtube.com/@creator',
  description: 'desc',
};

const youtubeEpisodeB: MediaLibraryItem = {
  ...youtubeEpisodeA,
  videoId: 2,
  canonicalTitle: 'Episode 2',
  youtubeVideoId: 'yt-2',
  videoUrl: 'https://www.youtube.com/watch?v=yt-2',
  videoTitle: 'Video 2',
  videoThumbnailUrl: 'https://i.ytimg.com/vi/yt-2/hqdefault.jpg',
  lastWatchedMs: 4_000,
};

const localVideo: MediaLibraryItem = {
  videoId: 3,
  canonicalTitle: 'Local Movie',
  totalSessions: 1,
  totalActiveMs: 5_000,
  totalCards: 0,
  totalTokensSeen: 40,
  lastWatchedMs: 2_000,
  hasCoverArt: 1,
  youtubeVideoId: null,
  videoUrl: null,
  videoTitle: null,
  videoThumbnailUrl: null,
  channelId: null,
  channelName: null,
  channelUrl: null,
  channelThumbnailUrl: null,
  uploaderId: null,
  uploaderUrl: null,
  description: null,
};

test('groupMediaLibraryItems groups youtube videos by channel and leaves local media standalone', () => {
  const groups = groupMediaLibraryItems([youtubeEpisodeA, localVideo, youtubeEpisodeB]);

  assert.equal(groups.length, 2);
  assert.equal(groups[0]?.title, 'Creator Name');
  assert.equal(groups[0]?.items.length, 2);
  assert.equal(groups[0]?.items[0]?.videoId, 2);
  assert.equal(groups[0]?.imageUrl, 'https://yt3.googleusercontent.com/channel-avatar=s88');
  assert.equal(groups[1]?.title, 'Local Movie');
  assert.equal(groups[1]?.items.length, 1);
});

test('groupMediaLibraryItems falls back to channel metadata when youtube channel id is missing', () => {
  const first = {
    ...youtubeEpisodeA,
    videoId: 20,
    youtubeVideoId: 'yt-20',
    videoUrl: 'https://www.youtube.com/watch?v=yt-20',
    channelId: null,
  };
  const second = {
    ...youtubeEpisodeB,
    videoId: 21,
    youtubeVideoId: 'yt-21',
    videoUrl: 'https://www.youtube.com/watch?v=yt-21',
    channelId: null,
  };

  const groups = groupMediaLibraryItems([first, second]);

  assert.equal(groups.length, 1);
  assert.equal(groups[0]?.title, 'Creator Name');
  assert.equal(groups[0]?.items.length, 2);
  assert.equal(groups[0]?.channelUrl, 'https://www.youtube.com/channel/UC123');
});

test('resolveMediaArtworkUrl prefers youtube thumbnails for video and channel images', () => {
  assert.equal(
    resolveMediaArtworkUrl(youtubeEpisodeA, 'video'),
    'https://i.ytimg.com/vi/yt-1/hqdefault.jpg',
  );
  assert.equal(
    resolveMediaArtworkUrl(youtubeEpisodeA, 'channel'),
    'https://yt3.googleusercontent.com/channel-avatar=s88',
  );
  assert.equal(resolveMediaArtworkUrl(localVideo, 'video'), null);
  assert.equal(resolveMediaArtworkUrl(localVideo, 'channel'), null);
});

test('resolveMediaArtworkUrl normalizes blank thumbnail urls to null', () => {
  const item = {
    videoThumbnailUrl: '   ',
    channelThumbnailUrl: '',
  };

  assert.equal(resolveMediaArtworkUrl(item, 'video'), null);
  assert.equal(resolveMediaArtworkUrl(item, 'channel'), null);
});

test('summarizeMediaLibraryGroups stays aligned with rendered group buckets', () => {
  const groups = groupMediaLibraryItems([youtubeEpisodeA, localVideo, youtubeEpisodeB]);
  const summary = summarizeMediaLibraryGroups(groups);

  assert.deepEqual(summary, {
    totalMs: 29_000,
    totalVideos: 3,
  });
});

test('groupMediaLibraryItems backfills missing group artwork from later items', () => {
  const first = {
    ...youtubeEpisodeA,
    videoId: 10,
    videoThumbnailUrl: null,
    channelThumbnailUrl: null,
  };
  const second = {
    ...youtubeEpisodeB,
    videoId: 11,
    channelThumbnailUrl: null,
  };

  const groups = groupMediaLibraryItems([first, second]);

  assert.equal(groups[0]?.imageUrl, second.videoThumbnailUrl);
});

test('CoverImage renders explicit remote artwork when src is provided', () => {
  const markup = renderToStaticMarkup(
    <CoverImage
      videoId={youtubeEpisodeA.videoId}
      title={youtubeEpisodeA.canonicalTitle}
      src={youtubeEpisodeA.videoThumbnailUrl}
      className="w-8 h-8"
    />,
  );

  assert.match(markup, /src="https:\/\/i\.ytimg\.com\/vi\/yt-1\/hqdefault\.jpg"/);
});

test('MediaCard uses the proxied cover endpoint instead of metadata artwork urls', () => {
  const markup = renderToStaticMarkup(<MediaCard item={youtubeEpisodeA} onClick={() => {}} />);

  assert.match(markup, /src="http:\/\/127\.0\.0\.1:6969\/api\/stats\/media\/1\/cover"/);
  assert.doesNotMatch(markup, /https:\/\/i\.ytimg\.com\/vi\/yt-1\/hqdefault\.jpg/);
});

test('resolveMediaCoverApiUrl appends retry tokens for late cover refreshes', () => {
  assert.equal(
    resolveMediaCoverApiUrl(youtubeEpisodeA.videoId, 2),
    'http://127.0.0.1:6969/api/stats/media/1/cover?coverRetry=2',
  );
});

test('MediaCard prefers youtube video title over canonical fallback url slug', () => {
  const markup = renderToStaticMarkup(<MediaCard item={youtubeEpisodeA} onClick={() => {}} />);

  assert.match(markup, />Video 1</);
  assert.match(markup, />Episode 1</);
});

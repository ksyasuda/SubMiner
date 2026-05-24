import assert from 'node:assert/strict';
import test from 'node:test';
import {
  isSameYoutubeMediaPath,
  isYoutubeMediaPath,
  isYoutubePlaybackActive,
  shouldUseCachedYoutubeParsedCues,
} from './youtube-playback';

test('isYoutubeMediaPath detects youtube watch and short urls', () => {
  assert.equal(isYoutubeMediaPath('https://www.youtube.com/watch?v=abc123'), true);
  assert.equal(isYoutubeMediaPath('https://m.youtube.com/watch?v=abc123'), true);
  assert.equal(isYoutubeMediaPath('https://youtu.be/abc123'), true);
  assert.equal(isYoutubeMediaPath('https://www.youtube-nocookie.com/embed/abc123'), true);
});

test('isYoutubeMediaPath ignores local files and non-youtube urls', () => {
  assert.equal(isYoutubeMediaPath('/tmp/video.mkv'), false);
  assert.equal(isYoutubeMediaPath('https://example.com/watch?v=abc123'), false);
  assert.equal(isYoutubeMediaPath('https://notyoutube.com/watch?v=abc123'), false);
  assert.equal(isYoutubeMediaPath('   '), false);
  assert.equal(isYoutubeMediaPath(null), false);
});

test('isYoutubePlaybackActive checks both current media and mpv video paths', () => {
  assert.equal(isYoutubePlaybackActive('/tmp/video.mkv', 'https://youtu.be/abc123'), true);
  assert.equal(isYoutubePlaybackActive('https://www.youtube.com/watch?v=abc123', null), true);
  assert.equal(isYoutubePlaybackActive('/tmp/video.mkv', '/tmp/video.mkv'), false);
});

test('isSameYoutubeMediaPath matches equivalent youtube urls by video id', () => {
  assert.equal(
    isSameYoutubeMediaPath('https://www.youtube.com/watch?v=abc123&t=30', 'https://youtu.be/abc123'),
    true,
  );
  assert.equal(
    isSameYoutubeMediaPath(
      'https://www.youtube.com/embed/abc123',
      'https://www.youtube-nocookie.com/embed/abc123',
    ),
    true,
  );
  assert.equal(
    isSameYoutubeMediaPath('https://www.youtube.com/watch?v=abc123', 'https://youtu.be/xyz789'),
    false,
  );
  assert.equal(isSameYoutubeMediaPath('/tmp/video.mkv', 'https://youtu.be/abc123'), false);
});

test('shouldUseCachedYoutubeParsedCues requires cached cues for the same youtube video', () => {
  assert.equal(
    shouldUseCachedYoutubeParsedCues({
      videoPath: 'https://www.youtube.com/watch?v=abc123&t=30',
      cachedMediaPath: 'https://youtu.be/abc123',
      cachedCueCount: 12,
    }),
    true,
  );
  assert.equal(
    shouldUseCachedYoutubeParsedCues({
      videoPath: 'https://www.youtube.com/watch?v=abc123',
      cachedMediaPath: 'https://youtu.be/abc123',
      cachedCueCount: 0,
    }),
    false,
  );
  assert.equal(
    shouldUseCachedYoutubeParsedCues({
      videoPath: 'https://www.youtube.com/watch?v=abc123',
      cachedMediaPath: 'https://youtu.be/other',
      cachedCueCount: 12,
    }),
    false,
  );
});

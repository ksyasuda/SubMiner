import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createStreamPlaybackMetadataStore,
  matchRequestedStreamPlaybackMetadata,
  toAnilistMediaGuess,
  toJimakuMediaInfo,
} from './stream-playback-metadata';
import type { AnimeStreamMetadata } from '../../anime-bridge/episode-metadata';

function metadata(overrides: Partial<AnimeStreamMetadata> = {}): AnimeStreamMetadata {
  return {
    mediaPath: 'http://127.0.0.1:41234/video/abc.m3u8',
    statsPath: 'animebrowser://9001/%2Fanime%2Fmushoku/%2Fwatch%2Fep-4',
    seriesTitle: 'Mushoku Tensei: Jobless Reincarnation',
    seasonNumber: 3,
    episodeNumber: 4,
    episodeTitle: null,
    displayTitle: 'Mushoku Tensei: Jobless Reincarnation S03E04',
    ...overrides,
  };
}

test('the store answers for the stream URL and for the stats path', () => {
  const store = createStreamPlaybackMetadataStore();
  const current = metadata();
  store.set(current);

  assert.equal(store.match(current.mediaPath), current);
  assert.equal(store.match(current.statsPath), current);
  assert.equal(store.match(`  ${current.mediaPath}  `), current);
});

test('the store stops answering once the player moves on', () => {
  const store = createStreamPlaybackMetadataStore();
  store.set(metadata());

  assert.equal(store.match('/home/user/Videos/local.mkv'), null);
  assert.equal(store.match(null), null);
  assert.equal(store.match(''), null);

  store.clear();
  assert.equal(store.match(metadata().mediaPath), null);
});

test('an explicit target path does not inherit the current stream metadata', () => {
  const store = createStreamPlaybackMetadataStore();
  const current = metadata();
  store.set(current);

  assert.equal(
    matchRequestedStreamPlaybackMetadata(store, '/home/user/Videos/other.mkv', current.mediaPath),
    null,
  );
  assert.equal(matchRequestedStreamPlaybackMetadata(store, null, current.mediaPath), current);
});

test('toJimakuMediaInfo prefills the modals with the source-reported fields', () => {
  assert.deepEqual(toJimakuMediaInfo(metadata()), {
    title: 'Mushoku Tensei: Jobless Reincarnation',
    season: 3,
    episode: 4,
    confidence: 'high',
    filename: 'Mushoku Tensei: Jobless Reincarnation S03E04',
    rawTitle: 'Mushoku Tensei: Jobless Reincarnation S03E04',
  });
});

test('toJimakuMediaInfo drops to low confidence when there is no episode', () => {
  const info = toJimakuMediaInfo(metadata({ episodeNumber: null, seasonNumber: null }));
  assert.equal(info.episode, null);
  assert.equal(info.confidence, 'low');
  assert.equal(info.title, 'Mushoku Tensei: Jobless Reincarnation');
});

test('toJimakuMediaInfo reports low confidence when a fractional episode is omitted', () => {
  const info = toJimakuMediaInfo(metadata({ episodeNumber: 6.5 }));
  assert.equal(info.episode, null);
  assert.equal(info.confidence, 'low');
});

test('toAnilistMediaGuess reports the stream fields verbatim', () => {
  assert.deepEqual(toAnilistMediaGuess(metadata()), {
    title: 'Mushoku Tensei: Jobless Reincarnation',
    season: 3,
    episode: 4,
    source: 'stream',
  });
});

test('toAnilistMediaGuess declines a special so the caller can fall back', () => {
  // AniList progress counts whole episodes; a 6.5 special is not one of them.
  assert.equal(toAnilistMediaGuess(metadata({ episodeNumber: 6.5 })), null);
  assert.equal(toAnilistMediaGuess(metadata({ episodeNumber: null })), null);
  assert.equal(toAnilistMediaGuess(metadata({ seriesTitle: '' })), null);
});

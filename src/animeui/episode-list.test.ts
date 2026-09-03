import assert from 'node:assert/strict';
import test from 'node:test';
import { nextPlaybackCue, playingEpisodeForAnime } from './episode-playback';

const selected = { sourceId: 'source.one', url: '/anime/one', title: 'One' };

test('playing episode sync applies only to the matching source and anime', () => {
  assert.equal(
    playingEpisodeForAnime(
      { sourceId: 'source.one', animeUrl: '/anime/one', episodeUrl: '/episode/3' },
      selected,
    ),
    '/episode/3',
  );
  assert.equal(
    playingEpisodeForAnime(
      { sourceId: 'source.two', animeUrl: '/anime/one', episodeUrl: '/episode/3' },
      selected,
    ),
    null,
  );
  assert.equal(
    playingEpisodeForAnime(
      { sourceId: 'source.one', animeUrl: '/anime/two', episodeUrl: '/episode/3' },
      selected,
    ),
    null,
  );
  assert.equal(playingEpisodeForAnime(null, selected), null);
});

test('live playback sync preserves a pending episode cue until resolution finishes', () => {
  const loading = { url: '/episode/4', state: 'loading' as const };

  assert.equal(
    nextPlaybackCue(
      { sourceId: 'source.one', animeUrl: '/anime/one', episodeUrl: '/episode/3' },
      selected,
      loading,
    ),
    loading,
  );
  assert.deepEqual(
    nextPlaybackCue(
      { sourceId: 'source.one', animeUrl: '/anime/one', episodeUrl: '/episode/3' },
      selected,
      { url: '/episode/2', state: 'playing' },
    ),
    { url: '/episode/3', state: 'playing' },
  );
});

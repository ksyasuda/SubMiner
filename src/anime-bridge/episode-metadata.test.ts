import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildAnimeStreamMetadata,
  buildAnimeStreamStatsPath,
  buildStreamDisplayTitle,
  splitEpisodeLabel,
  splitSeasonFromTitle,
} from './episode-metadata';

test('splitSeasonFromTitle pulls a trailing season marker off the title', () => {
  assert.deepEqual(splitSeasonFromTitle('Mushoku Tensei: Jobless Reincarnation Season 3'), {
    title: 'Mushoku Tensei: Jobless Reincarnation',
    season: 3,
  });
  assert.deepEqual(splitSeasonFromTitle('Spy x Family 2nd Season'), {
    title: 'Spy x Family',
    season: 2,
  });
  assert.deepEqual(splitSeasonFromTitle('Bocchi the Rock! S2'), {
    title: 'Bocchi the Rock!',
    season: 2,
  });
  assert.deepEqual(splitSeasonFromTitle('シャングリラ・フロンティア 第2期'), {
    title: 'シャングリラ・フロンティア',
    season: 2,
  });
});

test('splitSeasonFromTitle leaves a title without a trailing marker alone', () => {
  assert.deepEqual(splitSeasonFromTitle('My Teen Romantic Comedy SNAFU Climax!'), {
    title: 'My Teen Romantic Comedy SNAFU Climax!',
    season: null,
  });
  // "Season" inside the name is not a season marker.
  assert.deepEqual(splitSeasonFromTitle('A Season of Snow and Ash'), {
    title: 'A Season of Snow and Ash',
    season: null,
  });
  // Nothing would be left of the title, so the marker is not a marker.
  assert.deepEqual(splitSeasonFromTitle('Season 2'), { title: 'Season 2', season: null });
});

test('splitEpisodeLabel reads the number and the episode name', () => {
  assert.deepEqual(splitEpisodeLabel('Episode 4'), { number: 4, title: null });
  assert.deepEqual(splitEpisodeLabel('Episode 10: Gallantly, Shizuka Hiratsuka Moves Forward.'), {
    number: 10,
    title: 'Gallantly, Shizuka Hiratsuka Moves Forward.',
  });
  assert.deepEqual(splitEpisodeLabel('Ep. 7 - The Long Road'), {
    number: 7,
    title: 'The Long Road',
  });
  assert.deepEqual(splitEpisodeLabel('第12話 決戦'), { number: 12, title: '決戦' });
  assert.deepEqual(splitEpisodeLabel('5. Homecoming'), { number: 5, title: 'Homecoming' });
  assert.deepEqual(splitEpisodeLabel('13'), { number: 13, title: null });
  assert.deepEqual(splitEpisodeLabel('Episode 6.5'), { number: 6.5, title: null });
});

test('splitEpisodeLabel keeps a label that carries no number as a name', () => {
  assert.deepEqual(splitEpisodeLabel('Movie'), { number: null, title: 'Movie' });
  assert.deepEqual(splitEpisodeLabel('OVA - Beach Episode'), {
    number: null,
    title: 'OVA - Beach Episode',
  });
  assert.deepEqual(splitEpisodeLabel(''), { number: null, title: null });
});

test('buildStreamDisplayTitle emits a form guessit and the jimaku parser both read', () => {
  assert.equal(
    buildStreamDisplayTitle('Mushoku Tensei: Jobless Reincarnation', 3, 4, null),
    'Mushoku Tensei: Jobless Reincarnation S03E04',
  );
  assert.equal(
    buildStreamDisplayTitle('My Teen Romantic Comedy SNAFU Climax!', null, 10, 'Gallantly'),
    'My Teen Romantic Comedy SNAFU Climax! E10 - Gallantly',
  );
  assert.equal(buildStreamDisplayTitle('Some Movie', null, null, null), 'Some Movie');
});

test('buildAnimeStreamStatsPath is stable across playbacks of the same episode', () => {
  const first = buildAnimeStreamStatsPath('9001', '/anime/mushoku', '/watch/ep-4');
  const second = buildAnimeStreamStatsPath('9001', '/anime/mushoku', '/watch/ep-4');
  assert.equal(first, second);
  assert.notEqual(first, buildAnimeStreamStatsPath('9001', '/anime/mushoku', '/watch/ep-5'));
  assert.match(first, /^animebrowser:\/\//);
});

test('buildAnimeStreamMetadata resolves the browser strings into fields', () => {
  const metadata = buildAnimeStreamMetadata({
    sourceId: '9001',
    animeUrl: '/anime/mushoku',
    animeTitle: 'Mushoku Tensei: Jobless Reincarnation Season 3',
    episodeUrl: '/watch/ep-4',
    episodeName: 'Episode 4',
    episodeNumber: 4,
    mediaPath: 'http://127.0.0.1:41234/video/abc123.m3u8',
  });

  assert.equal(metadata.seriesTitle, 'Mushoku Tensei: Jobless Reincarnation');
  assert.equal(metadata.seasonNumber, 3);
  assert.equal(metadata.episodeNumber, 4);
  assert.equal(metadata.episodeTitle, null);
  assert.equal(metadata.displayTitle, 'Mushoku Tensei: Jobless Reincarnation S03E04');
  assert.equal(metadata.mediaPath, 'http://127.0.0.1:41234/video/abc123.m3u8');
  assert.equal(
    metadata.statsPath,
    buildAnimeStreamStatsPath('9001', '/anime/mushoku', '/watch/ep-4'),
  );
});

test('buildAnimeStreamMetadata prefers the extension episode number over the label', () => {
  const metadata = buildAnimeStreamMetadata({
    sourceId: '1',
    animeUrl: '/a',
    animeTitle: 'Show',
    episodeUrl: '/e',
    episodeName: 'Finale',
    episodeNumber: 24,
    mediaPath: 'http://host/x.m3u8',
  });
  assert.equal(metadata.episodeNumber, 24);
  assert.equal(metadata.episodeTitle, 'Finale');
  assert.equal(metadata.displayTitle, 'Show E24 - Finale');
});

test('buildAnimeStreamMetadata falls back to the label when the source reports no number', () => {
  const metadata = buildAnimeStreamMetadata({
    sourceId: '1',
    animeUrl: '/a',
    animeTitle: 'Show 2nd Season',
    episodeUrl: '/e',
    episodeName: 'Episode 3: Rain',
    episodeNumber: null,
    mediaPath: 'http://host/x.m3u8',
  });
  assert.equal(metadata.seasonNumber, 2);
  assert.equal(metadata.episodeNumber, 3);
  assert.equal(metadata.displayTitle, 'Show S02E03 - Rain');
});

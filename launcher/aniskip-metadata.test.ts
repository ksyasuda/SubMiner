import test from 'node:test';
import assert from 'node:assert/strict';
import {
  inferAniSkipMetadataForFile,
  buildSubminerScriptOpts,
  parseAniSkipGuessitJson,
} from './aniskip-metadata';

test('parseAniSkipGuessitJson extracts title season and episode', () => {
  const parsed = parseAniSkipGuessitJson(
    JSON.stringify({ title: 'My Show', season: 2, episode: 7 }),
    '/tmp/My.Show.S02E07.mkv',
  );
  assert.deepEqual(parsed, {
    title: 'My Show',
    season: 2,
    episode: 7,
    source: 'guessit',
  });
});

test('parseAniSkipGuessitJson prefers series over episode title', () => {
  const parsed = parseAniSkipGuessitJson(
    JSON.stringify({
      title: 'What Is This, a Picnic',
      series: 'Solo Leveling',
      season: 1,
      episode: 10,
    }),
    '/tmp/S01E10-What.Is.This,.a.Picnic-Bluray-1080p.mkv',
  );
  assert.deepEqual(parsed, {
    title: 'Solo Leveling',
    season: 1,
    episode: 10,
    source: 'guessit',
  });
});

test('inferAniSkipMetadataForFile falls back to filename title when guessit unavailable', () => {
  const parsed = inferAniSkipMetadataForFile('/tmp/Another_Show_-_03.mkv', {
    commandExists: () => false,
    runGuessit: () => null,
  });
  assert.equal(parsed.title.length > 0, true);
  assert.equal(parsed.source, 'fallback');
});

test('inferAniSkipMetadataForFile falls back to anime directory title when filename is episode-only', () => {
  const parsed = inferAniSkipMetadataForFile(
    '/truenas/jellyfin/anime/Solo Leveling/Season-1/S01E10-What.Is.This,.a.Picnic-Bluray-1080p.mkv',
    {
      commandExists: () => false,
      runGuessit: () => null,
    },
  );
  assert.equal(parsed.title, 'Solo Leveling');
  assert.equal(parsed.season, 1);
  assert.equal(parsed.episode, 10);
  assert.equal(parsed.source, 'fallback');
});

test('buildSubminerScriptOpts includes aniskip metadata fields', () => {
  const opts = buildSubminerScriptOpts('/tmp/SubMiner.AppImage', '/tmp/subminer.sock', {
    title: 'Frieren: Beyond Journey\'s End',
    season: 1,
    episode: 5,
    source: 'guessit',
  });
  assert.match(opts, /subminer-binary_path=\/tmp\/SubMiner\.AppImage/);
  assert.match(opts, /subminer-socket_path=\/tmp\/subminer\.sock/);
  assert.match(opts, /subminer-aniskip_title=Frieren: Beyond Journey's End/);
  assert.match(opts, /subminer-aniskip_season=1/);
  assert.match(opts, /subminer-aniskip_episode=5/);
});

import assert from 'node:assert/strict';
import test from 'node:test';
import { describeAnimeBrowserError } from './error-message';

test('missing bridge API explains incompatibility and preserves the method signature', () => {
  const message =
    "Anime bridge getVideoList failed (500). 'java.lang.Object eu.kanade.tachiyomi.animesource.online.AnimeHttpSource.getHosterList(eu.kanade.tachiyomi.animesource.model.SEpisode, kotlin.coroutines.Continuation)'";
  const error = describeAnimeBrowserError(message);
  assert.equal(error.title, 'Could not resolve the video');
  assert.match(error.explanation, /installed extension bridge does not provide/);
  assert.match(error.guidance, /If none are available/);
  assert.equal(error.details, message);
});

test('incomplete extension data is distinct from a missing bridge API', () => {
  const message =
    'Anime bridge getDetailsAnime failed (500). lateinit property url has not been initialized';
  const error = describeAnimeBrowserError(message);
  assert.equal(error.title, 'Could not load anime details');
  assert.match(error.explanation, /required field/);
  assert.equal(error.details, message);
});

test('unknown bridge failures do not claim an incompatibility', () => {
  for (const message of [
    'Anime bridge getEpisodeList failed (500). Unexpected response',
    'Anime bridge getEpisodeList failed (502).',
    'Anime bridge getEpisodeList failed: Cloudflare challenge',
  ]) {
    const error = describeAnimeBrowserError(message);
    assert.equal(error.title, 'Could not load episodes');
    assert.equal(error.explanation, 'The extension bridge could not complete this request.');
    assert.equal(error.details, message);
  }
});

test('playback and ordinary errors keep their own explanations', () => {
  const playback = describeAnimeBrowserError('mpv could not play this stream: loading failed');
  assert.equal(playback.title, 'Could not start playback');
  assert.match(playback.guidance, /fresh stream/);
  const other = describeAnimeBrowserError('Select a source first.');
  assert.equal(other.explanation, 'Select a source first.');
  assert.equal(other.details, '');
});

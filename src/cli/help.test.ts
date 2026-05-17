import test from 'node:test';
import assert from 'node:assert/strict';
import { printHelp } from './help';

test('printHelp includes configured texthooker port', () => {
  const original = console.log;
  let output = '';
  console.log = (value?: unknown) => {
    output += String(value);
  };

  try {
    printHelp(7777);
  } finally {
    console.log = original;
  }

  assert.match(output, /--help\s+Show this help/);
  assert.match(output, /default: 7777/);
  assert.match(output, /--launch-mpv.*Launch mpv with SubMiner defaults and exit/);
  assert.match(output, /--stats\s+Open the stats dashboard in your browser/);
  assert.match(output, /--open-browser\s+Open texthooker in your default browser/);
  assert.doesNotMatch(output, /--refresh-known-words/);
  assert.match(output, /--setup\s+Open first-run setup window/);
  assert.match(output, /--config\s+Open configuration window/);
  assert.match(output, /--anilist-status/);
  assert.match(output, /--anilist-retry-queue/);
  assert.match(output, /--dictionary/);
  assert.match(output, /--dictionary-target/);
  assert.match(output, /--jellyfin\s+Open Jellyfin setup window/);
  assert.match(output, /--jellyfin-login/);
  assert.match(output, /--jellyfin-subtitles/);
  assert.match(output, /--jellyfin-play/);
});

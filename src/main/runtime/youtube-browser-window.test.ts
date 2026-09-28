import assert from 'node:assert/strict';
import test from 'node:test';
import { buildYoutubeBrowserUserAgent } from './youtube-browser-window';

test('buildYoutubeBrowserUserAgent strips Electron and app tokens from the user agent', () => {
  assert.equal(
    buildYoutubeBrowserUserAgent(
      'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) SubMiner/0.19.6 Chrome/140.0.7339.41 Electron/42.6.0 Safari/537.36',
    ),
    'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.7339.41 Safari/537.36',
  );
});

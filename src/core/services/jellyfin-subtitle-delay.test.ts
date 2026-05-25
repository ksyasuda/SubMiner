import assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import test from 'node:test';
import { loadJellyfinSubtitleDelay, saveJellyfinSubtitleDelay } from './jellyfin-subtitle-delay';

function statePath(name: string): string {
  return path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-jellyfin-delay-')), name);
}

test('jellyfin subtitle delay store saves and loads delay by item and stream', () => {
  const filePath = statePath('delays.json');

  assert.equal(
    saveJellyfinSubtitleDelay({
      filePath,
      itemId: 'episode-1',
      streamIndex: 3,
      delaySeconds: 1.25,
    }),
    true,
  );

  assert.equal(loadJellyfinSubtitleDelay({ filePath, itemId: 'episode-1', streamIndex: 3 }), 1.25);
  assert.equal(loadJellyfinSubtitleDelay({ filePath, itemId: 'episode-1', streamIndex: 4 }), null);
});

test('jellyfin subtitle delay store preserves other stream delays when updating one stream', () => {
  const filePath = statePath('delays.json');

  saveJellyfinSubtitleDelay({ filePath, itemId: 'episode-1', streamIndex: 3, delaySeconds: 1.25 });
  saveJellyfinSubtitleDelay({ filePath, itemId: 'episode-1', streamIndex: 4, delaySeconds: -0.5 });
  saveJellyfinSubtitleDelay({ filePath, itemId: 'episode-1', streamIndex: 3, delaySeconds: 2 });

  assert.equal(loadJellyfinSubtitleDelay({ filePath, itemId: 'episode-1', streamIndex: 3 }), 2);
  assert.equal(loadJellyfinSubtitleDelay({ filePath, itemId: 'episode-1', streamIndex: 4 }), -0.5);
});

test('jellyfin subtitle delay store ignores invalid files and values', () => {
  const filePath = statePath('delays.json');
  fs.writeFileSync(filePath, '{');

  assert.equal(loadJellyfinSubtitleDelay({ filePath, itemId: 'episode-1', streamIndex: 3 }), null);
  assert.equal(
    saveJellyfinSubtitleDelay({
      filePath,
      itemId: 'episode-1',
      streamIndex: 3,
      delaySeconds: Number.NaN,
    }),
    false,
  );
});

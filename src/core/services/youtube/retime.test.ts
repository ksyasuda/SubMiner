import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { retimeYoutubeSubtitle } from './retime';

test('retimeYoutubeSubtitle uses the downloaded subtitle path as-is', async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-youtube-retime-'));
  try {
    const primaryPath = path.join(root, 'primary.vtt');
    const referencePath = path.join(root, 'reference.vtt');
    fs.writeFileSync(primaryPath, 'WEBVTT\n', 'utf8');
    fs.writeFileSync(referencePath, 'WEBVTT\n', 'utf8');

    const result = await retimeYoutubeSubtitle({
      primaryPath,
      secondaryPath: referencePath,
    });

    assert.equal(result.ok, true);
    assert.equal(result.strategy, 'none');
    assert.equal(result.path, primaryPath);
    assert.equal(result.message, 'Using downloaded subtitle as-is (no automatic retime enabled)');
    assert.equal(fs.readFileSync(result.path, 'utf8'), 'WEBVTT\n');
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});

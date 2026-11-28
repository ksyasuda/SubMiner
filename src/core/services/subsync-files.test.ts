import test from 'node:test';
import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { cleanupTemporaryFile, pruneRetainedSubsyncTempDirs } from './subsync-files';

test('synced embedded-track temp dirs keep only the loaded file and go away once mpv drops it', () => {
  const isolatedTmp = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-files-test-'));
  try {
    // Left behind by a crashed session two days ago.
    const staleDir = fs.mkdtempSync(path.join(isolatedTmp, 'subminer-subsync-'));
    const twoDaysAgo = (Date.now() - 2 * 24 * 60 * 60 * 1000) / 1000;
    fs.utimesSync(staleDir, twoDaysAgo, twoDaysAgo);

    // Extraction of an embedded target plus the retimed copy mpv loaded.
    const runDir = fs.mkdtempSync(path.join(isolatedTmp, 'subminer-subsync-'));
    const extractedPath = path.join(runDir, 'track_2.ass');
    const retimedPath = path.join(runDir, 'track_2_retimed.ass');
    fs.writeFileSync(extractedPath, 'original');
    fs.writeFileSync(retimedPath, 'retimed');

    cleanupTemporaryFile({ path: extractedPath, temporary: true });
    assert.deepEqual(fs.readdirSync(runDir), ['track_2_retimed.ass']);

    // An old dir whose file mpv still plays survives the sweep.
    const playingDir = fs.mkdtempSync(path.join(isolatedTmp, 'subminer-subsync-'));
    const playingPath = path.join(playingDir, 'track_4_retimed.srt');
    fs.writeFileSync(playingPath, 'still loaded');
    fs.utimesSync(playingDir, twoDaysAgo, twoDaysAgo);

    pruneRetainedSubsyncTempDirs([retimedPath, playingPath], { tmpRoot: isolatedTmp });
    assert.equal(fs.existsSync(runDir), true, 'mpv still lists the retimed file');
    assert.equal(fs.existsSync(staleDir), false, 'stale dirs from old sessions are swept');
    assert.equal(fs.existsSync(playingDir), true, 'stale but referenced dirs are kept');

    pruneRetainedSubsyncTempDirs(['/media/ep02.ja.srt'], { tmpRoot: isolatedTmp });
    assert.equal(fs.existsSync(runDir), false, 'dropped once mpv no longer lists it');
  } finally {
    fs.rmSync(isolatedTmp, { recursive: true, force: true });
  }
});

test('cleanupTemporaryFile removes a reference extraction and its dir entirely', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-subsync-'));
  const extractedPath = path.join(dir, 'track_3.srt');
  fs.writeFileSync(extractedPath, 'reference');

  cleanupTemporaryFile({ path: extractedPath, temporary: true });
  assert.equal(fs.existsSync(dir), false);
});

import test from 'node:test';
import assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import {
  discoverSidecarReferenceFiles,
  rankReferenceCandidates,
  type EvaluatedSubtitleReference,
} from './reference-discovery';

test('discoverSidecarReferenceFiles finds unloaded sidecar subtitle files and skips target/retimed/loaded', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-sidecars-'));
  try {
    const videoPath = path.join(dir, 'Episode 01.mkv');
    const targetPath = path.join(dir, 'Episode 01.ja.srt');
    const loadedEngPath = path.join(dir, 'Episode 01.en.srt');
    const unloadedAssPath = path.join(dir, 'Episode 01.eng.full.ass');
    const retimedPath = path.join(dir, 'Episode 01.ja_retimed.srt');
    const otherEpPath = path.join(dir, 'Episode 02.en.srt');

    for (const p of [
      videoPath,
      targetPath,
      loadedEngPath,
      unloadedAssPath,
      retimedPath,
      otherEpPath,
    ]) {
      fs.writeFileSync(p, 'test');
    }

    const found = discoverSidecarReferenceFiles(videoPath, [targetPath, loadedEngPath]);
    assert.deepEqual(
      found.map((f) => f.filename),
      ['Episode 01.eng.full.ass'],
    );
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('rankReferenceCandidates prefers audio-aligned unloaded track over misaligned secondary track', () => {
  const misalignedSecondary: EvaluatedSubtitleReference = {
    key: 'track:2',
    kind: 'track',
    trackId: 2,
    filePath: null,
    label: 'External #2 - eng (secondary)',
    isSecondary: true,
    isEmbedded: false,
    isEnglish: true,
    isFullDialogueHint: true,
    isSignsOnly: false,
    dialogueCount: 300,
    signCount: 10,
    audioStatus: 'misaligned',
    audioOffsetMs: 2400,
    audioConfidence: 0.92,
  };

  const alignedUnloaded: EvaluatedSubtitleReference = {
    key: 'track:4',
    kind: 'track',
    trackId: 4,
    filePath: null,
    label: 'Internal #4 - eng - Full Dialogue',
    isSecondary: false,
    isEmbedded: true,
    isEnglish: true,
    isFullDialogueHint: true,
    isSignsOnly: false,
    dialogueCount: 295,
    signCount: 20,
    audioStatus: 'aligned',
    audioOffsetMs: 40,
    audioConfidence: 0.9,
  };

  const signsOnly: EvaluatedSubtitleReference = {
    key: 'track:3',
    kind: 'track',
    trackId: 3,
    filePath: null,
    label: 'Internal #3 - eng - Signs & Songs',
    isSecondary: false,
    isEmbedded: true,
    isEnglish: true,
    isFullDialogueHint: false,
    isSignsOnly: true,
    dialogueCount: 12,
    signCount: 180,
    audioStatus: 'unverified',
    audioOffsetMs: null,
    audioConfidence: 0,
  };

  const ranked = rankReferenceCandidates([misalignedSecondary, signsOnly, alignedUnloaded]);
  assert.deepEqual(
    ranked.map((r) => r.key),
    ['track:4', 'track:3', 'track:2'],
  );
});

import test from 'node:test';
import assert from 'node:assert/strict';
import {
  alignNoSplit,
  alignSubtitleFileToReferenceSpans,
  evaluateReferenceAlignmentToAudio,
  evaluateReferenceAlignmentToSampledAudio,
  joinSpans,
  scoreStanding,
  CALIBRATION_CLEAN,
  wrappedCoverage,
  type SampledAudioWindow,
} from './alass-align';
import {
  formatSrtTimestamp,
  parseSubtitleFileContent,
  type SubtitleTimeSpan,
} from './subtitle-lines';

function generateDialogueRhythm(
  seedInit: bigint,
  step: bigint,
  untilSec = 1300,
  gapScale = 4.0,
): SubtitleTimeSpan[] {
  const out: SubtitleTimeSpan[] = [];
  let seed = seedInit;
  const mask = (1n << 64n) - 1n;
  let t = 5.0;
  while (t < untilSec) {
    seed = (seed * 6364136223846793005n + step) & mask;
    const len = 0.8 + (Number(seed >> 33n) / Number(1n << 31n)) * 3.0;
    const gap = 0.4 + (Number((seed >> 13n) & 0xffffn) / 65536.0) * gapScale;
    out.push({ start: t, end: t + len });
    t += len + gap;
  }
  return out;
}

function spansToSrtFile(spans: readonly SubtitleTimeSpan[]): string {
  return spans
    .map(
      (s, i) =>
        `${i + 1}\n${formatSrtTimestamp(Math.max(0, s.start))} --> ${formatSrtTimestamp(Math.max(0.05, s.end))}\nセリフ ${i + 1}\n`,
    )
    .join('\n');
}

test('joinSpans merges overlapping spans so stacked karaoke lines count once', () => {
  const merged = joinSpans([
    { start: 10, end: 15 },
    { start: 10, end: 15 },
    { start: 12, end: 18 },
    { start: 20, end: 22 },
  ]);
  assert.deepEqual(merged, [
    { start: 10, end: 18 },
    { start: 20, end: 22 },
  ]);
});

test('alignNoSplit finds exact large and small offsets', () => {
  const ref = generateDialogueRhythm(7n, 1442695040888963407n);
  const shifted = ref.map((s) => ({ start: s.start + 42.35, end: s.end + 42.35 }));
  const delta = alignNoSplit(joinSpans(ref), shifted);
  assert.ok(Math.abs(delta + 42.35) <= 0.05, `expected ~-42.35s, got ${delta}`);
});

test('alignNoSplit lands on the exact overlap peak, not one bin early', () => {
  const cue = [{ start: 10, end: 11 }];
  assert.equal(alignNoSplit(cue, cue, { centerShiftSec: 0.5, radiusSec: 2 }), 0);
  // Overlap plateau spans [-2.5s, 0s]; the sweep reports its first bin.
  assert.equal(
    alignNoSplit([{ start: 10, end: 10.5 }], [{ start: 10, end: 13 }], {
      centerShiftSec: 0,
      radiusSec: 5,
    }),
    -2.5,
  );
});

test('alignSubtitleFileToReferenceSpans recovers fixed shift and ignores stacked songs in reference', () => {
  const dialogue = generateDialogueRhythm(7n, 1442695040888963407n);
  const refWithSong: SubtitleTimeSpan[] = [...dialogue];
  // Simulate 8 simultaneous stacked karaoke lines during an ending song.
  for (let i = 0; i < 8; i += 1) {
    refWithSong.push({ start: 1310, end: 1400 });
  }

  const targetSpans = dialogue.map((s) => ({
    start: s.start + 2.5 - 0.08,
    end: s.end + 2.5 + 0.15,
  }));
  const parsedTarget = parseSubtitleFileContent(spansToSrtFile(targetSpans), 'ep01.ja.srt');

  const result = alignSubtitleFileToReferenceSpans(parsedTarget, refWithSong, {
    calibration: 'clean',
    lengthSec: 1420,
  });

  assert.equal(result.confident, true);
  assert.ok(result.confidence >= 0.8, `expected high confidence, got ${result.confidence}`);
  assert.ok(Math.abs(result.offsetMs + 2500) <= 120, `expected ~-2500ms, got ${result.offsetMs}`);
  assert.equal(result.segments, 1);
});

test('alignSubtitleFileToReferenceSpans corrects PAL 25fps speed-up automatically', () => {
  const ref = generateDialogueRhythm(7n, 1442695040888963407n);
  const palRatio = 23.976 / 25.0;
  const palSpans = ref.map((s) => ({
    start: s.start / palRatio + 1.2,
    end: s.end / palRatio + 1.2,
  }));
  const parsedTarget = parseSubtitleFileContent(spansToSrtFile(palSpans), 'pal.ja.srt');

  const result = alignSubtitleFileToReferenceSpans(parsedTarget, ref, {
    calibration: 'clean',
    lengthSec: 1320,
  });

  assert.equal(result.confident, true);
  assert.ok(
    Math.abs(result.ratio - palRatio) < 1e-5,
    `expected ratio ${palRatio}, got ${result.ratio}`,
  );
});

for (const lengthSec of [1300, 2700]) {
  test(`alignSubtitleFileToReferenceSpans handles TV broadcast ad breaks in pieces (${lengthSec}s)`, () => {
    const ref = generateDialogueRhythm(7n, 1442695040888963407n, lengthSec);
    const broadcastSpans = ref.map((s) => {
      const adOffset = s.start < lengthSec / 3 ? 30 : s.start < (2 * lengthSec) / 3 ? 120 : 195;
      return { start: s.start + adOffset, end: s.end + adOffset };
    });
    const parsedTarget = parseSubtitleFileContent(spansToSrtFile(broadcastSpans), 'tv.ja.srt');

    const result = alignSubtitleFileToReferenceSpans(parsedTarget, ref, {
      calibration: 'clean',
      lengthSec: lengthSec + 20,
      allowSplits: true,
    });

    assert.equal(result.cut, true);
    assert.equal(result.segments, 3);
    assert.equal(result.confident, true);

    const retimed = parseSubtitleFileContent(result.rewrittenContent, 'tv.ja.srt');
    for (let i = 0; i < ref.length; i += 1) {
      assert.ok(
        Math.abs(retimed.events[i]!.start - ref[i]!.start) < 0.25,
        `line ${i} misaligned: got ${retimed.events[i]!.start}, expected ${ref[i]!.start}`,
      );
    }
  });
}

test('alignSubtitleFileToReferenceSpans refines adjacent scene cuts without moving neighbors', () => {
  const ref = generateDialogueRhythm(7n, 1442695040888963407n, 1300, 12);
  // Lines 34..49 straddle two 16-line blocks, so both of their boundaries are cuts.
  const target = ref.map((s, i) =>
    i >= 34 && i <= 49 ? { start: s.start + 6, end: s.end + 6 } : s,
  );
  const result = alignSubtitleFileToReferenceSpans(
    parseSubtitleFileContent(spansToSrtFile(target), 'cuts.ja.srt'),
    ref,
    { calibration: 'clean', lengthSec: 1320, allowSplits: true },
  );

  assert.equal(result.segments, 3);
  const retimed = parseSubtitleFileContent(result.rewrittenContent, 'cuts.ja.srt');
  for (let i = 0; i < ref.length; i += 1) {
    assert.ok(
      Math.abs(retimed.events[i]!.start - ref[i]!.start) < 0.15,
      `line ${i} misaligned: got ${retimed.events[i]!.start}, expected ${ref[i]!.start}`,
    );
  }
});

test('alignSubtitleFileToReferenceSpans is not confident when only a recap section matches', () => {
  const ref = generateDialogueRhythm(7n, 1442695040888963407n);
  const otherEpisode = generateDialogueRhythm(12345n, 0x9e3779b97f4a7c15n);
  // First quarter shared (a recap), the rest from a different episode.
  const target = [
    ...ref.filter((s) => s.start < 325),
    ...otherEpisode.filter((s) => s.start >= 330),
  ].map((s) => ({ start: s.start + 3, end: s.end + 3 }));

  const result = alignSubtitleFileToReferenceSpans(
    parseSubtitleFileContent(spansToSrtFile(target), 'recap.ja.srt'),
    ref,
    { calibration: 'clean', lengthSec: 1320 },
  );

  assert.equal(result.confident, false);
  assert.ok(
    result.confidence < 0.65,
    `expected sub-threshold confidence, got ${result.confidence}`,
  );
});

test('alignSubtitleFileToReferenceSpans rejects an unrelated episode with low confidence', () => {
  const ref = generateDialogueRhythm(7n, 1442695040888963407n);
  const otherEpisode = generateDialogueRhythm(12345n, 0x9e3779b97f4a7c15n);
  const parsedTarget = parseSubtitleFileContent(spansToSrtFile(otherEpisode), 'wrong.ja.srt');

  const result = alignSubtitleFileToReferenceSpans(parsedTarget, ref, {
    calibration: 'clean',
    lengthSec: 1320,
  });

  assert.equal(result.confident, false);
  assert.ok(result.confidence < 0.4, `expected low confidence, got ${result.confidence}`);
});

test('evaluateReferenceAlignmentToAudio distinguishes aligned vs shifted English reference tracks', () => {
  const audioSpeech = generateDialogueRhythm(7n, 1442695040888963407n);
  // Well-aligned English track (leads audio by ~0.1s).
  const goodEng = audioSpeech.map((s) => ({ start: Math.max(0, s.start - 0.1), end: s.end + 0.1 }));
  // Misaligned English track (shifted +2.8s off audio).
  const badEng = audioSpeech.map((s) => ({ start: s.start + 2.8, end: s.end + 2.8 }));

  const goodEval = evaluateReferenceAlignmentToAudio(goodEng, audioSpeech, {
    audioCalibration: 'noisy',
    lengthSec: 1320,
  });
  const badEval = evaluateReferenceAlignmentToAudio(badEng, audioSpeech, {
    audioCalibration: 'noisy',
    lengthSec: 1320,
  });

  assert.equal(goodEval.status, 'aligned');
  assert.ok(
    Math.abs(goodEval.offsetMs) <= 300,
    `expected near-0ms offset, got ${goodEval.offsetMs}`,
  );

  assert.equal(badEval.status, 'misaligned');
  assert.ok(
    Math.abs(badEval.offsetMs + 2900) <= 300,
    `expected ~-2900ms offset, got ${badEval.offsetMs}`,
  );
});

test('evaluateReferenceAlignmentToSampledAudio verifies aligned vs misaligned refs and marks huge shifts unverified', () => {
  const fullSpeech = generateDialogueRhythm(7n, 1442695040888963407n);
  const windowStarts = [160, 460, 780, 1080];
  const windowDur = 60;
  const windows: SampledAudioWindow[] = windowStarts.map((startSec) => ({
    startSec,
    durationSec: windowDur,
    speechSpans: fullSpeech
      .filter((s) => s.end > startSec && s.start < startSec + windowDur)
      .map((s) => ({
        start: Math.max(0, s.start - startSec),
        end: Math.min(windowDur, s.end - startSec),
      })),
  }));

  const alignedRef = fullSpeech.map((s) => ({
    start: Math.max(0, s.start - 0.1),
    end: s.end + 0.1,
  }));
  const shiftedRef = fullSpeech.map((s) => ({ start: s.start + 3.2, end: s.end + 3.2 }));
  const hugeShiftRef = fullSpeech.map((s) => ({ start: s.start + 45.0, end: s.end + 45.0 }));

  const alignedEval = evaluateReferenceAlignmentToSampledAudio(alignedRef, windows);
  assert.equal(alignedEval.status, 'aligned');
  assert.ok(Math.abs(alignedEval.offsetMs) <= 350, `got ${alignedEval.offsetMs}`);

  const shiftedEval = evaluateReferenceAlignmentToSampledAudio(shiftedRef, windows);
  assert.equal(shiftedEval.status, 'misaligned');
  assert.ok(Math.abs(shiftedEval.offsetMs + 3300) <= 350, `got ${shiftedEval.offsetMs}`);

  // A 45s shift is outside the ±12s sample window search radius -> returns 'unverified' so caller escalates to full audio.
  const hugeEval = evaluateReferenceAlignmentToSampledAudio(hugeShiftRef, windows);
  assert.equal(hugeEval.status, 'unverified');
});

test('wrappedCoverage counts both fragments of a cue crossing the wrap boundary', () => {
  const line = [{ start: 54, end: 58 }];
  const ref = [
    { start: 0, end: 3 },
    { start: 59, end: 60 },
  ];
  // Shifting [54, 58] by +5s in a 60s window wraps to [59, 60] (1s) + [0, 3] (3s) = 4s / 4s = 1.0.
  assert.equal(wrappedCoverage(line, 4, ref, 5, 60), 1);
  // Also works for shifts larger than the window length and negative shifts.
  assert.equal(wrappedCoverage(line, 4, ref, 65, 60), 1);
  assert.equal(
    wrappedCoverage(
      [{ start: 1, end: 5 }],
      4,
      [
        { start: 0, end: 2 },
        { start: 58, end: 60 },
      ],
      -3,
      60,
    ),
    1,
  );
});

test('alignSubtitleFileToReferenceSpans rejects broadcast pieces that collapse onto the same reference section', () => {
  const actPattern = generateDialogueRhythm(7n, 1442695040888963407n).filter((s) => s.end < 120);
  // Two identical target acts separated by a long pause within the standing shift range,
  // while the reference only contains one matching section.
  const repeatedTarget = [
    ...actPattern,
    ...actPattern.map((s) => ({ start: s.start + 150, end: s.end + 150 })),
  ];
  const parsedTarget = parseSubtitleFileContent(spansToSrtFile(repeatedTarget), 'dup-acts.ja.srt');

  const result = alignSubtitleFileToReferenceSpans(parsedTarget, actPattern, {
    calibration: 'clean',
    lengthSec: 320,
    allowSplits: true,
  });

  assert.equal(result.cut, false);
  assert.equal(result.confident, false);
});

test('scoreStanding caps a weak-quarter fit unless every quarter agrees on line starts', () => {
  const strongCoverage = { above: 14, alone: 5, weakestPart: 0.5 };
  // One quarter does not fit and starts do not agree everywhere: below the 0.65 auto bar.
  assert.ok(scoreStanding({ ...strongCoverage, starts: 0 }, CALIBRATION_CLEAN) <= 0.6);
  // Every quarter agrees line for line (starts is the weakest quarter's z): not a partial fit.
  assert.ok(scoreStanding({ ...strongCoverage, starts: 9 }, CALIBRATION_CLEAN) > 0.65);
});

import test from 'node:test';
import assert from 'node:assert/strict';
import { alignSubtitleFileToReferenceSpans } from './alass-align';
import { alignSubtitleOffThread } from './align-worker';
import {
  formatSrtTimestamp,
  parseSubtitleFileContent,
  type SubtitleTimeSpan,
} from './subtitle-lines';

// Deterministic dialogue-like spans (same LCG as alass-align.test.ts).
function rhythm(untilSec: number): SubtitleTimeSpan[] {
  const out: SubtitleTimeSpan[] = [];
  let seed = 7n;
  const mask = (1n << 64n) - 1n;
  let t = 5;
  while (t < untilSec) {
    seed = (seed * 6364136223846793005n + 1442695040888963407n) & mask;
    const len = 0.8 + (Number(seed >> 33n) / Number(1n << 31n)) * 3;
    const gap = 0.4 + (Number((seed >> 13n) & 0xffffn) / 65536) * 4;
    out.push({ start: t, end: t + len });
    t += len + gap;
  }
  return out;
}

function srtShiftedBy(spans: readonly SubtitleTimeSpan[], shiftSec: number): string {
  return spans
    .map(
      (s, i) =>
        `${i + 1}\n${formatSrtTimestamp(s.start + shiftSec)} --> ${formatSrtTimestamp(s.end + shiftSec)}\nline ${i + 1}\n`,
    )
    .join('\n');
}

test('alignSubtitleOffThread returns the same alignment as the in-thread aligner', async () => {
  const reference = rhythm(1300);
  const target = parseSubtitleFileContent(srtShiftedBy(reference, 2.5), 'target.srt');

  const offThread = await alignSubtitleOffThread(undefined, target, reference, {});
  assert.deepEqual(offThread, alignSubtitleFileToReferenceSpans(target, reference, {}));
  assert.equal(offThread.offsetMs, -2500);
});

test('aborting alignSubtitleOffThread stops a long alignment without waiting for it', async () => {
  // A 2.5 h film takes most of a second to align; an abort can only win if it runs off-thread.
  const reference = rhythm(9000);
  const target = parseSubtitleFileContent(srtShiftedBy(reference, 4), 'movie.srt');
  const startedAt = Date.now();
  const inThreadMs = (() => {
    alignSubtitleFileToReferenceSpans(target, reference, {});
    return Date.now() - startedAt;
  })();

  const controller = new AbortController();
  const abortedAt = Date.now();
  setTimeout(() => controller.abort(new Error('media changed')), 20);
  await assert.rejects(
    alignSubtitleOffThread(controller.signal, target, reference, {}),
    /media changed/,
  );
  assert.ok(
    Date.now() - abortedAt < inThreadMs,
    `abort took ${Date.now() - abortedAt}ms, in-thread alignment takes ${inThreadMs}ms`,
  );

  // The next call starts a fresh worker.
  const after = await alignSubtitleOffThread(undefined, target, reference, {});
  assert.equal(after.offsetMs, -4000);
});

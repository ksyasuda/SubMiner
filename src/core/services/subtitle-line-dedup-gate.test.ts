import assert from 'node:assert/strict';
import test from 'node:test';
import { createSubtitleLineDedupGate } from './subtitle-line-dedup-gate';
import type { SubtitleCue } from '../../types';

function karaokeFrames(text: string, start: number, frames: number, frameSeconds: number) {
  return Array.from({ length: frames }, (_, index) => ({
    text,
    startSec: start + index * frameSeconds,
    endSec: start + (index + 1) * frameSeconds,
  }));
}

test('parsed cues drop the frames the sidebar already collapsed', () => {
  // What `mergeDuplicateCues` leaves behind for a karaoke run: one cue over the run.
  const cues: SubtitleCue[] = [
    { startTime: 10, endTime: 14, text: '飛び上がる' },
    { startTime: 14, endTime: 16, text: 'もしも' },
  ];
  const gate = createSubtitleLineDedupGate({ getParsedCues: () => cues });

  const recorded = karaokeFrames('飛び上がる', 10, 40, 0.1).filter((sample) =>
    gate.shouldRecord(sample),
  );

  assert.equal(recorded.length, 1);
  assert.equal(recorded[0]!.startSec, 10);
  assert.equal(gate.shouldRecord({ text: 'もしも', startSec: 14, endSec: 16 }), true);
});

test('parsed cues keep separate lines that merely repeat', () => {
  const cues: SubtitleCue[] = [
    { startTime: 3, endTime: 3.4, text: 'えっ' },
    { startTime: 3.4, endTime: 3.9, text: 'えっ' },
    { startTime: 3.9, endTime: 4.5, text: 'えっ' },
  ];
  const gate = createSubtitleLineDedupGate({ getParsedCues: () => cues });

  const recorded = cues.filter((cue) =>
    gate.shouldRecord({ text: cue.text, startSec: cue.startTime, endSec: cue.endTime }),
  );

  assert.equal(recorded.length, 3);
});

test('a line whose timing does not match any cue still records', () => {
  // A shifted track, an embedded sub nobody parsed: no match, no drop.
  const cues: SubtitleCue[] = [{ startTime: 10, endTime: 14, text: '飛び上がる' }];
  const gate = createSubtitleLineDedupGate({ getParsedCues: () => cues });

  assert.equal(gate.shouldRecord({ text: '飛び上がる', startSec: 42, endSec: 44 }), true);
});

test('without parsed cues a long run of identical short frames stops recording', () => {
  const gate = createSubtitleLineDedupGate({ getParsedCues: () => null });

  const recorded = karaokeFrames('ひとしずく', 0, 200, 0.04).filter((sample) =>
    gate.shouldRecord(sample),
  );

  assert.equal(recorded.length, 4);
});

test('without parsed cues ordinary repeated dialogue keeps recording', () => {
  const gate = createSubtitleLineDedupGate({ getParsedCues: () => null });

  // Six contiguous `えっ`, each held for a normal beat rather than an animation frame.
  const recorded = karaokeFrames('えっ', 0, 6, 0.6).filter((sample) => gate.shouldRecord(sample));

  assert.equal(recorded.length, 6);
});

test('the same event offered twice does not advance the run', () => {
  const gate = createSubtitleLineDedupGate({ getParsedCues: () => null });

  // mpv fires the timing handler once for `sub-start` and once for `sub-end`.
  for (let i = 0; i < 8; i += 1) {
    assert.equal(gate.shouldRecord({ text: '待って', startSec: 5, endSec: 5.05 }), true);
  }
});

test('a gap between frames starts a new run', () => {
  const gate = createSubtitleLineDedupGate({ getParsedCues: () => null });

  const first = karaokeFrames('もし', 0, 6, 0.04).filter((sample) => gate.shouldRecord(sample));
  const second = karaokeFrames('もし', 30, 6, 0.04).filter((sample) => gate.shouldRecord(sample));

  assert.equal(first.length, 4);
  assert.equal(second.length, 4);
});

test('reset forgets the streaming run', () => {
  const gate = createSubtitleLineDedupGate({ getParsedCues: () => null });

  karaokeFrames('もし', 0, 20, 0.04).forEach((sample) => gate.shouldRecord(sample));
  gate.reset();

  assert.equal(gate.shouldRecord({ text: 'もし', startSec: 0.8, endSec: 0.84 }), true);
});

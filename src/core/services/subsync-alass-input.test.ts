import test from 'node:test';
import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
  convertSubtitleForAlass,
  formatCuesAsSrt,
  needsAlassConversion,
  parseTimedCues,
} from './subsync-alass-input';

const VTT = [
  'WEBVTT',
  '',
  'NOTE this comment is not a cue',
  '',
  'cue-1',
  '00:00:01.500 --> 00:00:03.250 line:90% align:center',
  '<v Geese>しっかし',
  'まさかお前が',
  '',
  '00:01:02.000 --> 00:01:04.000',
  'ゼニスも きっと幸せじゃろう',
  '',
].join('\n');

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'alass-input-test-'));
}

test('parseTimedCues keeps cue payloads verbatim and skips headers', () => {
  const cues = parseTimedCues(VTT);
  assert.equal(cues.length, 2);
  assert.equal(cues[0]!.start, 1.5);
  assert.equal(cues[0]!.end, 3.25);
  assert.equal(cues[0]!.text, '<v Geese>しっかし\nまさかお前が');
  assert.equal(cues[1]!.start, 62);
  assert.equal(cues[1]!.text, 'ゼニスも きっと幸せじゃろう');
});

test('parseTimedCues reads VTT timestamps without an hours field', () => {
  const cues = parseTimedCues('WEBVTT\n\n01:02.000 --> 01:03.000\nhello\n');
  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.start, 62);
  assert.equal(cues[0]!.end, 63);
});

test('formatCuesAsSrt writes numbered SubRip blocks', () => {
  assert.equal(
    formatCuesAsSrt(parseTimedCues(VTT)),
    '1\n00:00:01,500 --> 00:00:03,250\n<v Geese>しっかし\nまさかお前が\n\n' +
      '2\n00:01:02,000 --> 00:01:04,000\nゼニスも きっと幸せじゃろう\n',
  );
});

test('needsAlassConversion targets VTT by extension and by content', () => {
  assert.equal(needsAlassConversion('/tmp/track-0.vtt', 'WEBVTT\n'), true);
  assert.equal(needsAlassConversion('/tmp/track-0.srt', 'WEBVTT\n'), true);
  assert.equal(needsAlassConversion('/tmp/track-0.ttml', '<?xml version="1.0"?>'), true);
  assert.equal(
    needsAlassConversion('/tmp/track-0.srt', '1\n00:00:01,000 --> 00:00:02,000\n'),
    false,
  );
  assert.equal(needsAlassConversion('/tmp/track-0.ass', '[Script Info]\n'), false);
});

test('convertSubtitleForAlass rewrites a VTT track as SRT', () => {
  const dir = makeTempDir();
  const source = path.join(dir, 'track-0.vtt');
  fs.writeFileSync(source, VTT);

  const converted = convertSubtitleForAlass(source);
  assert.notEqual(converted, null);
  assert.equal(converted!.temporary, true);
  assert.equal(path.extname(converted!.path), '.srt');
  assert.match(fs.readFileSync(converted!.path, 'utf8'), /00:00:01,500 --> 00:00:03,250/);

  fs.rmSync(dir, { recursive: true, force: true });
  fs.rmSync(path.dirname(converted!.path), { recursive: true, force: true });
});

test('convertSubtitleForAlass converts a VTT payload hiding behind an .srt name', () => {
  const dir = makeTempDir();
  const source = path.join(dir, 'remote_track.srt');
  fs.writeFileSync(source, VTT);

  const converted = convertSubtitleForAlass(source);
  assert.notEqual(converted, null);
  assert.notEqual(converted!.path, source);
  assert.equal(fs.readFileSync(converted!.path, 'utf8').startsWith('1\n'), true);

  fs.rmSync(dir, { recursive: true, force: true });
  fs.rmSync(path.dirname(converted!.path), { recursive: true, force: true });
});

test('convertSubtitleForAlass leaves an SRT file alone', () => {
  const dir = makeTempDir();
  const source = path.join(dir, 'track-0.srt');
  fs.writeFileSync(source, '1\n00:00:01,000 --> 00:00:02,000\nhello\n');

  assert.equal(convertSubtitleForAlass(source), null);
  fs.rmSync(dir, { recursive: true, force: true });
});

test('convertSubtitleForAlass reports a file with no readable timings', () => {
  const dir = makeTempDir();
  const source = path.join(dir, 'track-0.vtt');
  fs.writeFileSync(source, 'WEBVTT\n\nnot a cue at all\n');

  assert.throws(() => convertSubtitleForAlass(source), /Could not read subtitle timings/);
  fs.rmSync(dir, { recursive: true, force: true });
});

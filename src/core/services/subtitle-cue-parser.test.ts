import assert from 'node:assert/strict';
import test from 'node:test';
import { parseSrtCues, parseAssCues, parseSubtitleCues } from './subtitle-cue-parser';
import type { SubtitleCue } from '../../types';

test('parseSrtCues parses basic SRT content', () => {
  const content = [
    '1',
    '00:00:01,000 --> 00:00:04,000',
    'こんにちは',
    '',
    '2',
    '00:00:05,000 --> 00:00:08,500',
    '元気ですか',
    '',
  ].join('\n');

  const cues = parseSrtCues(content);

  assert.equal(cues.length, 2);
  assert.equal(cues[0]!.startTime, 1.0);
  assert.equal(cues[0]!.endTime, 4.0);
  assert.equal(cues[0]!.text, 'こんにちは');
  assert.equal(cues[1]!.startTime, 5.0);
  assert.equal(cues[1]!.endTime, 8.5);
  assert.equal(cues[1]!.text, '元気ですか');
});

test('parseSrtCues handles multi-line subtitle text', () => {
  const content = ['1', '00:01:00,000 --> 00:01:05,000', 'これは', 'テストです', ''].join('\n');

  const cues = parseSrtCues(content);

  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.text, 'これは\nテストです');
});

test('parseSrtCues handles hours in timestamps', () => {
  const content = ['1', '01:30:00,000 --> 01:30:05,000', 'テスト', ''].join('\n');

  const cues = parseSrtCues(content);

  assert.equal(cues[0]!.startTime, 5400.0);
  assert.equal(cues[0]!.endTime, 5405.0);
});

test('parseSrtCues handles VTT-style dot separator', () => {
  const content = ['1', '00:00:01.000 --> 00:00:04.000', 'VTTスタイル', ''].join('\n');

  const cues = parseSrtCues(content);

  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.startTime, 1.0);
});

test('parseSrtCues returns empty array for empty content', () => {
  assert.deepEqual(parseSrtCues(''), []);
  assert.deepEqual(parseSrtCues('  \n\n  '), []);
});

test('parseSrtCues skips malformed timing lines gracefully', () => {
  const content = [
    '1',
    'NOT A TIMING LINE',
    'テスト',
    '',
    '2',
    '00:00:01,000 --> 00:00:02,000',
    '有効',
    '',
  ].join('\n');

  const cues = parseSrtCues(content);

  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.text, '有効');
});

test('parseAssCues parses basic ASS dialogue lines', () => {
  const content = [
    '[Script Info]',
    'Title: Test',
    '',
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:04.00,Default,,0,0,0,,こんにちは',
    'Dialogue: 0,0:00:05.00,0:00:08.50,Default,,0,0,0,,元気ですか',
  ].join('\n');

  const cues = parseAssCues(content);

  assert.equal(cues.length, 2);
  assert.equal(cues[0]!.startTime, 1.0);
  assert.equal(cues[0]!.endTime, 4.0);
  assert.equal(cues[0]!.text, 'こんにちは');
  assert.equal(cues[1]!.startTime, 5.0);
  assert.equal(cues[1]!.endTime, 8.5);
  assert.equal(cues[1]!.text, '元気ですか');
});

test('parseAssCues strips override tags from text', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:04.00,Default,,0,0,0,,{\\b1}太字{\\b0}テスト',
  ].join('\n');

  const cues = parseAssCues(content);

  assert.equal(cues[0]!.text, '太字テスト');
});

test('parseAssCues handles text containing commas', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:04.00,Default,,0,0,0,,はい、そうです、ね',
  ].join('\n');

  const cues = parseAssCues(content);

  assert.equal(cues[0]!.text, 'はい、そうです、ね');
});

test('parseAssCues handles \\N line breaks', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:04.00,Default,,0,0,0,,一行目\\N二行目',
  ].join('\n');

  const cues = parseAssCues(content);

  assert.equal(cues[0]!.text, '一行目\\N二行目');
});

test('parseAssCues returns empty for content without Events section', () => {
  const content = ['[Script Info]', 'Title: Test'].join('\n');

  assert.deepEqual(parseAssCues(content), []);
});

test('parseAssCues skips Comment lines', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Comment: 0,0:00:01.00,0:00:04.00,Default,,0,0,0,,これはコメント',
    'Dialogue: 0,0:00:05.00,0:00:08.00,Default,,0,0,0,,これは字幕',
  ].join('\n');

  const cues = parseAssCues(content);

  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.text, 'これは字幕');
});

test('parseAssCues handles hour timestamps', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,1:30:00.00,1:30:05.00,Default,,0,0,0,,テスト',
  ].join('\n');

  const cues = parseAssCues(content);

  assert.equal(cues[0]!.startTime, 5400.0);
  assert.equal(cues[0]!.endTime, 5405.0);
});

test('parseAssCues respects dynamic field ordering from the Format row', () => {
  const content = [
    '[Events]',
    'Format: Layer, Style, Start, End, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,Default,0:00:01.00,0:00:04.00,,0,0,0,,順番が違う',
  ].join('\n');

  const cues = parseAssCues(content);

  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.startTime, 1.0);
  assert.equal(cues[0]!.endTime, 4.0);
  assert.equal(cues[0]!.text, '順番が違う');
});

test('parseSubtitleCues auto-detects SRT format', () => {
  const content = ['1', '00:00:01,000 --> 00:00:04,000', 'SRTテスト', ''].join('\n');

  const cues = parseSubtitleCues(content, 'test.srt');
  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.text, 'SRTテスト');
});

test('parseSubtitleCues auto-detects ASS format', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:04.00,Default,,0,0,0,,ASSテスト',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');
  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.text, 'ASSテスト');
});

test('parseSubtitleCues auto-detects VTT format', () => {
  const content = ['1', '00:00:01.000 --> 00:00:04.000', 'VTTテスト', ''].join('\n');

  const cues = parseSubtitleCues(content, 'test.vtt');
  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.text, 'VTTテスト');
});

test('parseSubtitleCues returns empty for unknown format', () => {
  assert.deepEqual(parseSubtitleCues('random content', 'test.xyz'), []);
});

test('parseSubtitleCues returns cues sorted by start time', () => {
  const content = [
    '1',
    '00:00:10,000 --> 00:00:14,000',
    '二番目',
    '',
    '2',
    '00:00:01,000 --> 00:00:04,000',
    '一番目',
    '',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.srt');
  assert.equal(cues[0]!.text, '一番目');
  assert.equal(cues[1]!.text, '二番目');
});

test('parseSubtitleCues detects subtitle formats from remote URLs', () => {
  const assContent = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:02.00,Default,,0,0,0,,URLテスト',
  ].join('\n');

  const cues = parseSubtitleCues(assContent, 'https://host/subs.ass?lang=ja#track');

  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.text, 'URLテスト');
});

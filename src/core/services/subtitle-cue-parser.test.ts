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

test('parseSrtCues preserves lines that only resemble malformed ASS controls', () => {
  const content = ['1', '00:01:00,000 --> 00:01:05,000', '\\', '{\\fr0', ''].join('\n');

  assert.equal(parseSrtCues(content)[0]?.text, '\\\n{\\fr0');
});

test('parseSrtCues strips HTML-like markup while preserving line breaks', () => {
  const content = [
    '1',
    '00:01:00,000 --> 00:01:05,000',
    '<font color="japanese">これは</font>',
    '<font color="japanese">テストです</font>',
    '',
  ].join('\n');

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

test('parseSubtitleCues strips complete brace blocks from SRT and VTT text', () => {
  const content = ['1', '00:00:01,000 --> 00:00:02,000', '彼は{謎}と言った', ''].join('\n');

  for (const filename of ['test.srt', 'test.vtt']) {
    const cues = parseSubtitleCues(content, filename);

    assert.equal(cues.length, 1, filename);
    assert.equal(cues[0]!.text, '彼はと言った', filename);
  }
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

test('parseAssCues decodes \\N line breaks into real newlines', () => {
  // ASS is decoded once, here at ingestion, so cue text matches what mpv hands over for
  // the same line played live.
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:04.00,Default,,0,0,0,,一行目\\N二行目',
  ].join('\n');

  const cues = parseAssCues(content);

  assert.equal(cues[0]!.text, '一行目\n二行目');
});

test('parseAssCues strips HTML-like markup while preserving ASS line breaks', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:04.00,Default,,0,0,0,,<font color="japanese">一行目</font>\\N<font color="japanese">二行目</font>',
  ].join('\n');

  const cues = parseAssCues(content);

  assert.equal(cues[0]!.text, '一行目\n二行目');
});

test('parseAssCues drops vector drawing runs enabled by \\p', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 1,0:00:01.00,0:00:04.00,Default,,0,0,0,,{\\an5\\pos(730,1042)\\p1\\blur1}m 20 0 b 10 0 0 10 0 20 b 0 31 10 40 20 40 {\\p0}',
    'Dialogue: 0,0:00:05.00,0:00:08.00,Default,,0,0,0,,これは字幕',
  ].join('\n');

  const cues = parseAssCues(content);

  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.text, 'これは字幕');
});

test('parseAssCues keeps text that follows a \\p0 reset on the same line', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:04.00,Default,,0,0,0,,{\\p1}m 0 0 l 10 10{\\p0}本文{\\p1}m 5 5 l 6 6{\\p0}続き',
  ].join('\n');

  const cues = parseAssCues(content);

  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.text, '本文続き');
});

test('parseAssCues leaves \\pos untouched when no drawing mode is active', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:04.00,Default,,0,0,0,,{\\pos(960,1068)\\bord3}位置指定',
  ].join('\n');

  const cues = parseAssCues(content);

  assert.equal(cues[0]!.text, '位置指定');
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

test('parseSubtitleCues collapses per-frame karaoke duplicates into one cue', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:01.05,OP_JP,,0,0,0,,{\\clip(m 1 1)}過ぎ去ってしまう瞬間を',
    'Dialogue: 0,0:00:01.05,0:00:01.09,OP_JP,,0,0,0,,{\\clip(m 2 2)}過ぎ去ってしまう瞬間を',
    'Dialogue: 0,0:00:01.09,0:00:03.55,OP_JP,,0,0,0,,{\\clip(m 3 3)}過ぎ去ってしまう瞬間を',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');

  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.startTime, 1.0);
  assert.equal(cues[0]!.endTime, 3.55);
  assert.equal(cues[0]!.text, '過ぎ去ってしまう瞬間を');
});

test('parseSubtitleCues collapses long full-line color phases', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 1,0:03:49.75,0:03:51.21,OPJP,,0,0,0,,{\\blur0.6\\c&H312D38&\\4c&HFFFFFF&}ちゃんと目を{\\4c&HD590FF&}合わせてよ',
    'Dialogue: 1,0:03:51.21,0:03:52.25,OPJP,,0,0,0,,{\\blur0.6\\4c&H312D38&\\c&HFFFFFF&}ちゃんと目を{\\4c&HD590FF&}合わせてよ',
  ].join('\n');

  assert.deepEqual(parseSubtitleCues(content, 'test.ass'), [
    {
      startTime: 229.75,
      endTime: 232.25,
      text: 'ちゃんと目を合わせてよ',
    },
  ]);
});

test('parseSubtitleCues keeps ordinary repeated dialogue separate', () => {
  // A single restyle tag on a repeated line is how ordinary dialogue gets decorated;
  // it is not phase evidence, whatever the line length.
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 1,0:00:01.00,0:00:02.00,OPJP,,0,0,0,,{\\c&H111111&}歌詞',
    'Dialogue: 1,0:00:02.00,0:00:03.00,OPJP,,0,0,0,,{\\c&H222222&}歌詞',
    'Dialogue: 1,0:00:04.00,0:00:05.00,OPJP,,0,0,0,,{\\c&H333333&}別の歌詞',
    'Dialogue: 1,0:00:05.00,0:00:06.00,OPJP,,0,0,0,,{\\c&H444444&}別の歌詞',
    'Dialogue: 8,0:00:07.00,0:00:08.00,Text - JP,,0,0,0,,えっ？',
    'Dialogue: 8,0:00:08.00,0:00:09.00,Text - JP,,0,0,0,,えっ？',
  ].join('\n');

  assert.deepEqual(parseSubtitleCues(content, 'test.ass'), [
    { startTime: 1, endTime: 2, text: '歌詞' },
    { startTime: 2, endTime: 3, text: '歌詞' },
    { startTime: 4, endTime: 5, text: '別の歌詞' },
    { startTime: 5, endTime: 6, text: '別の歌詞' },
    { startTime: 7, endTime: 8, text: 'えっ？' },
    { startTime: 8, endTime: 9, text: 'えっ？' },
  ]);
});

test('parseSubtitleCues keeps separately positioned temporal signs separate', () => {
  // Two flush signs with the same text but different \move paths are separate authored
  // occurrences, not phases of one redraw: temporal evidence alone must not merge them.
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:02.00,Sign,,0,0,0,,{\\move(100,100,200,100)}立入禁止',
    'Dialogue: 0,0:00:02.00,0:00:03.00,Sign,,0,0,0,,{\\move(500,400,600,400)}立入禁止',
  ].join('\n');

  assert.deepEqual(parseSubtitleCues(content, 'test.ass'), [
    { startTime: 1, endTime: 2, text: '立入禁止' },
    { startTime: 2, endTime: 3, text: '立入禁止' },
  ]);
});

test('parseSubtitleCues keeps richly styled ordinary repeats separate', () => {
  // Blur plus a changing color is still an ordinary restyle. Phase redraws are
  // recognized by the color/highlight boundary moving *within* the line, which these
  // leading-block-only events do not have.
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:02.00,Dial,,0,0,0,,{\\blur0.4\\c&H111111&}待ってよ',
    'Dialogue: 0,0:00:02.00,0:00:03.00,Dial,,0,0,0,,{\\blur0.4\\c&H222222&}待ってよ',
  ].join('\n');

  assert.deepEqual(parseSubtitleCues(content, 'test.ass'), [
    { startTime: 1, endTime: 2, text: '待ってよ' },
    { startTime: 2, endTime: 3, text: '待ってよ' },
  ]);
});

test('parseSubtitleCues keeps canonical metadata when an identical plain cue exists', () => {
  // A plain dialogue line can share exact timing and text with a recovered canonical
  // cue from another style. The canonical copy must win the exact-duplicate collapse,
  // or the live overlay loses the marker it substitutes on.
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:05.00,0:00:08.00,Plain,,0,0,0,,ライン',
    'Comment: 0,0:00:05.00,0:00:08.00,OP,,0,0,0,,ライン',
    'Dialogue: 0,0:00:05.00,0:00:05.04,OP,,0,0,0,,{\\pos(1,1)\\clip(m 1 1)}ライン',
    'Dialogue: 0,0:00:05.04,0:00:05.08,OP,,0,0,0,,{\\pos(1,1)\\clip(m 2 2)}ライン',
    'Dialogue: 0,0:00:05.08,0:00:08.00,OP,,0,0,0,,{\\pos(1,1)\\clip(m 3 3)}ライン',
  ].join('\n');

  assert.deepEqual(parseSubtitleCues(content, 'test.ass'), [
    {
      startTime: 5,
      endTime: 8,
      text: 'ライン',
      source: 'canonical-ass',
      animationStartTime: 5,
      animationEndTime: 8,
    },
  ]);
});

test('parseSubtitleCues keeps short styled repeats separate even with richer styling', () => {
  // Two ordinary えっ lines restyled with different colors are two utterances, not two
  // phases of one lyric: short text never satisfies the changing-override evidence path.
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:02.00,Dial,,0,0,0,,{\\blur0.4\\c&H111111&}えっ',
    'Dialogue: 0,0:00:02.00,0:00:03.00,Dial,,0,0,0,,{\\blur0.4\\c&H222222&}えっ',
  ].join('\n');

  assert.deepEqual(parseSubtitleCues(content, 'test.ass'), [
    { startTime: 1, endTime: 2, text: 'えっ' },
    { startTime: 2, endTime: 3, text: 'えっ' },
  ]);
});

test('parseSubtitleCues keeps back-to-back plain dialogue repeats separate', () => {
  // Several characters greeting in turn: distinct utterances that happen to abut.
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:04:05.67,0:04:06.82,Dial_JP,,0,0,0,,おはよう',
    'Dialogue: 0,0:04:06.82,0:04:07.56,Dial_JP,,0,0,0,,おはよう',
    'Dialogue: 0,0:04:07.56,0:04:08.78,Dial_JP,,0,0,0,,おはよう',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');

  assert.equal(cues.length, 3);
  assert.equal(cues[0]!.endTime, 246.82);
  assert.equal(cues[2]!.startTime, 247.56);
});

test('parseSubtitleCues collapses exact duplicate cues even without effect tags', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:04.00,Default,,0,0,0,,重なった行',
    'Dialogue: 1,0:00:01.00,0:00:04.00,Default,,0,0,0,,重なった行',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');

  assert.equal(cues.length, 1);
});

test('parseSubtitleCues replaces generated glyph animation with its timed canonical comment', () => {
  // Aegisub automation commonly keeps the authored lyric as a Comment and emits
  // multiple moving Dialogue layers for every glyph. This mirrors the MyGO ED script:
  // three entrance copies followed by three exit copies for each character.
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Comment: 0,0:00:01.20,0:00:03.80,ED_JP,,0,0,0,,{\\fad(480,480)}今　手にある',
    'Dialogue: 0,0:00:00.80,0:00:01.50,ED_JP,,0,0,0,,{\\move(10,20,100,200)\\t(0,600,\\fscx100)}今',
    'Dialogue: 0,0:00:00.80,0:00:01.50,ED_JP,,0,0,0,,{\\move(30,40,100,200)\\t(0,600,\\fscx100)}今',
    'Dialogue: 0,0:00:00.80,0:00:01.50,ED_JP,,0,0,0,,{\\move(50,60,100,200)\\t(0,600,\\fscx100)}今',
    'Dialogue: 1,0:00:01.40,0:00:04.20,ED_JP,,0,0,0,,{\\move(100,200,20,30)\\t(2000,2600,\\blur20)}今',
    'Dialogue: 1,0:00:01.40,0:00:04.20,ED_JP,,0,0,0,,{\\move(100,200,40,50)\\t(2000,2600,\\blur20)}今',
    'Dialogue: 1,0:00:01.40,0:00:04.20,ED_JP,,0,0,0,,{\\move(100,200,60,70)\\t(2000,2600,\\blur20)}今',
    'Dialogue: 0,0:00:00.86,0:00:01.56,ED_JP,,0,0,0,,{\\move(10,20,140,200)\\t(0,600,\\fscx100)}手',
    'Dialogue: 0,0:00:00.86,0:00:01.56,ED_JP,,0,0,0,,{\\move(30,40,140,200)\\t(0,600,\\fscx100)}手',
    'Dialogue: 0,0:00:00.86,0:00:01.56,ED_JP,,0,0,0,,{\\move(50,60,140,200)\\t(0,600,\\fscx100)}手',
    'Dialogue: 1,0:00:01.46,0:00:04.26,ED_JP,,0,0,0,,{\\move(140,200,20,30)\\t(2000,2600,\\blur20)}手',
    'Dialogue: 1,0:00:01.46,0:00:04.26,ED_JP,,0,0,0,,{\\move(140,200,40,50)\\t(2000,2600,\\blur20)}手',
    'Dialogue: 1,0:00:01.46,0:00:04.26,ED_JP,,0,0,0,,{\\move(140,200,60,70)\\t(2000,2600,\\blur20)}手',
    'Dialogue: 0,0:00:00.92,0:00:01.62,ED_JP,,0,0,0,,{\\move(10,20,180,200)\\t(0,600,\\fscx100)}にある',
    'Dialogue: 0,0:00:00.92,0:00:01.62,ED_JP,,0,0,0,,{\\move(30,40,180,200)\\t(0,600,\\fscx100)}にある',
    'Dialogue: 0,0:00:00.92,0:00:01.62,ED_JP,,0,0,0,,{\\move(50,60,180,200)\\t(0,600,\\fscx100)}にある',
    'Dialogue: 1,0:00:01.52,0:00:04.32,ED_JP,,0,0,0,,{\\move(180,200,20,30)\\t(2000,2600,\\blur20)}にある',
    'Dialogue: 1,0:00:01.52,0:00:04.32,ED_JP,,0,0,0,,{\\move(180,200,40,50)\\t(2000,2600,\\blur20)}にある',
    'Dialogue: 1,0:00:01.52,0:00:04.32,ED_JP,,0,0,0,,{\\move(180,200,60,70)\\t(2000,2600,\\blur20)}にある',
    'Dialogue: 0,0:00:06.00,0:00:08.00,Dial_JP,,0,0,0,,普通の会話',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');

  assert.deepEqual(cues, [
    {
      startTime: 1.2,
      endTime: 3.8,
      text: '今　手にある',
      source: 'canonical-ass',
      // Entrance frames start before and exit frames end after the authored timing.
      animationStartTime: 0.8,
      animationEndTime: 4.32,
    },
    { startTime: 6, endTime: 8, text: '普通の会話' },
  ]);
});

test('parseSubtitleCues recovers a full Dialogue line surrounding generated fragments', () => {
  // Some scripts do not retain the authored line as a Comment. Instead, brief entrance
  // and exit events contain the complete line around a long run of generated syllables.
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 1,0:00:01.00,0:00:01.15,ED Romaji,,0,0,0,fx,{\\move(100,40,60,40)}toki yo ugokidase',
    'Dialogue: 1,0:00:01.15,0:00:03.00,ED Romaji,,0,0,0,fx,{\\t(0,300,\\c&HFFFFFF&)}to',
    'Dialogue: 1,0:00:01.15,0:00:03.00,ED Romaji,,0,0,0,fx,{\\t(300,500,\\c&HFFFFFF&)}ki',
    'Dialogue: 1,0:00:01.15,0:00:03.00,ED Romaji,,0,0,0,fx,{\\t(500,700,\\c&HFFFFFF&)}yo',
    'Dialogue: 1,0:00:01.15,0:00:03.00,ED Romaji,,0,0,0,fx,{\\t(700,900,\\c&HFFFFFF&)}u',
    'Dialogue: 1,0:00:01.15,0:00:03.00,ED Romaji,,0,0,0,fx,{\\t(900,1100,\\c&HFFFFFF&)}go',
    'Dialogue: 1,0:00:01.15,0:00:03.00,ED Romaji,,0,0,0,fx,{\\t(1100,1300,\\c&HFFFFFF&)}ki',
    'Dialogue: 1,0:00:01.15,0:00:03.00,ED Romaji,,0,0,0,fx,{\\t(1300,1500,\\c&HFFFFFF&)}da',
    'Dialogue: 1,0:00:01.15,0:00:03.00,ED Romaji,,0,0,0,fx,{\\t(1500,1800,\\c&HFFFFFF&)}se',
    'Dialogue: 1,0:00:03.00,0:00:03.15,ED Romaji,,0,0,0,fx,{\\move(60,40,20,40)}toki yo ugokidase',
    'Dialogue: 0,0:00:06.00,0:00:08.00,Default,,0,0,0,,Ordinary dialogue',
  ].join('\n');

  assert.deepEqual(parseSubtitleCues(content, 'test.ass'), [
    {
      startTime: 1,
      endTime: 3.15,
      text: 'toki yo ugokidase',
      source: 'canonical-ass',
      animationStartTime: 1,
      animationEndTime: 3.15,
    },
    { startTime: 6, endTime: 8, text: 'Ordinary dialogue' },
  ]);
});

test('parseSubtitleCues does not promote a short animated fragment as a complete line', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 1,0:00:01.00,0:00:03.00,OP English,,0,0,0,,{\\pos(100,100)\\t(0,100,\\fscx120)}my',
    'Dialogue: 2,0:00:01.00,0:00:03.00,OP English,,0,0,0,,{\\pos(100,100)\\t(0,100,\\fscx120)}my',
    'Dialogue: 1,0:00:01.00,0:00:03.00,OP English,,0,0,0,,{\\pos(100,100)\\t(0,100,\\fscx120)}m',
    'Dialogue: 2,0:00:01.00,0:00:03.00,OP English,,0,0,0,,{\\pos(100,100)\\t(0,100,\\fscx120)}m',
    'Dialogue: 1,0:00:01.00,0:00:03.00,OP English,,0,0,0,,{\\pos(120,100)\\t(20,120,\\fscx120)}y',
    'Dialogue: 2,0:00:01.00,0:00:03.00,OP English,,0,0,0,,{\\pos(120,100)\\t(20,120,\\fscx120)}y',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');

  assert.equal(
    cues.some((cue) => cue.source === 'canonical-ass'),
    false,
  );
});

test('parseSubtitleCues keeps short animated English dialogue as separate cues', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:02.00,English Dialogue,,0,0,0,,{\\t(0,100,\\fscx110)}Hi',
    'Dialogue: 0,0:00:02.00,0:00:03.00,English Dialogue,,0,0,0,,{\\t(0,100,\\fscx110)}No',
  ].join('\n');

  assert.deepEqual(parseSubtitleCues(content, 'test.ass'), [
    { startTime: 1, endTime: 2, text: 'Hi' },
    { startTime: 2, endTime: 3, text: 'No' },
  ]);
});

test('parseSubtitleCues does not reconstruct an already canonical English cue', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Comment: 0,0:00:01.00,0:00:03.00,OP English,,0,0,0,,{\\move(100,100,120,100)}POOF',
    'Dialogue: 0,0:00:01.00,0:00:01.04,OP English,,0,0,0,,{\\pos(100,100)\\clip(m 1 1)}POOF',
    'Dialogue: 0,0:00:01.04,0:00:01.08,OP English,,0,0,0,,{\\pos(100,100)\\clip(m 2 2)}POOF',
    'Dialogue: 0,0:00:01.08,0:00:03.00,OP English,,0,0,0,,{\\pos(100,100)\\clip(m 3 3)}POOF',
  ].join('\n');

  assert.deepEqual(parseSubtitleCues(content, 'test.ass'), [
    {
      startTime: 1,
      endTime: 3,
      text: 'POOF',
      source: 'canonical-ass',
      animationStartTime: 1,
      animationEndTime: 3,
    },
  ]);
});

test('parseSubtitleCues reconstructs a short positioned fragment without a lyric style name', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:03.00,Karaoke,,0,0,0,,{\\pos(100,100)\\t(0,100,\\fscx110)}Oh',
    'Dialogue: 1,0:00:01.00,0:00:03.00,Karaoke,,0,0,0,,{\\pos(100,100)\\t(0,100,\\fscx110)}Oh',
  ].join('\n');

  assert.deepEqual(parseSubtitleCues(content, 'test.ass'), [
    {
      startTime: 1,
      endTime: 3,
      text: 'Oh',
      source: 'reconstructed-ass',
      animationStartTime: 1,
      animationEndTime: 3,
      assStyle: 'Karaoke',
    },
  ]);
});

test('parseSubtitleCues ignores timed comments without a matching animated dialogue cluster', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Comment: 0,0:00:01.00,0:00:03.00,Dial_JP,,0,0,0,,編集メモ',
    'Comment: 0,0:00:04.00,0:00:06.00,Dial_JP,,0,0,0,,別案の字幕',
    'Dialogue: 0,0:00:01.00,0:00:03.00,Dial_JP,,0,0,0,,通常の字幕',
    'Dialogue: 0,0:00:04.00,0:00:06.00,Dial_JP,,0,0,0,,別案の字幕',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');

  assert.deepEqual(cues, [
    { startTime: 1, endTime: 3, text: '通常の字幕' },
    { startTime: 4, endTime: 6, text: '別案の字幕' },
  ]);
});

test('parseAssCues returns recovered canonical cues in chronological order', () => {
  // Recovery appends recovered cues after surviving dialogue; the bare parseAssCues
  // export must still come back time-ordered.
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:06.00,0:00:08.00,Dial,,0,0,0,,あとのセリフ',
    'Comment: 0,0:00:01.20,0:00:03.80,OP,,0,0,0,,雨が上がっても',
    'Dialogue: 0,0:00:01.20,0:00:01.24,OP,,0,0,0,,{\\pos(1,1)\\clip(m 1 1)}雨が上がっても',
    'Dialogue: 0,0:00:01.24,0:00:01.28,OP,,0,0,0,,{\\pos(1,1)\\clip(m 2 2)}雨が上がっても',
    'Dialogue: 0,0:00:01.28,0:00:03.80,OP,,0,0,0,,{\\pos(1,1)\\clip(m 3 3)}雨が上がっても',
  ].join('\n');

  assert.deepEqual(
    parseAssCues(content).map((cue) => cue.startTime),
    [1.2, 6],
  );
});

test('parseSubtitleCues withdraws a recovery whose owner is claimed by a later candidate', () => {
  // The exit boundary event appears first in the file and recovers a canonical cue from
  // its own small cluster. The entrance candidate then proves that exit event was a
  // generated frame of the full animation; the earlier recovery is a duplicate of the
  // same authored line and must not survive alongside it.
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 1,0:00:14.00,0:00:14.20,ED,,0,0,0,,{\\move(100,200,20,30)}ABCDEFGH',
    'Dialogue: 1,0:00:13.50,0:00:14.50,ED,,0,0,0,,{\\t(0,300,\\c&HFFFFFF&)}ABC',
    'Dialogue: 1,0:00:13.50,0:00:14.50,ED,,0,0,0,,{\\t(300,600,\\c&HFFFFFF&)}DEF',
    'Dialogue: 1,0:00:13.50,0:00:14.50,ED,,0,0,0,,{\\t(600,900,\\c&HFFFFFF&)}GH',
    'Dialogue: 0,0:00:10.00,0:00:10.20,ED,,0,0,0,,{\\move(10,20,100,200)}ABCDEFGH',
    'Dialogue: 0,0:00:10.00,0:00:12.00,ED,,0,0,0,,{\\t(0,300,\\fscx100)}ABC',
    'Dialogue: 0,0:00:10.00,0:00:12.00,ED,,0,0,0,,{\\t(300,600,\\fscx100)}DEF',
    'Dialogue: 0,0:00:10.00,0:00:13.40,ED,,0,0,0,,{\\t(600,900,\\fscx100)}GH',
  ].join('\n');

  assert.deepEqual(parseSubtitleCues(content, 'test.ass'), [
    {
      startTime: 10,
      endTime: 14.2,
      text: 'ABCDEFGH',
      source: 'canonical-ass',
      animationStartTime: 10,
      animationEndTime: 14.2,
    },
  ]);
});

test('parseSubtitleCues recovers canonical comments from generated clip frames', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Comment: 0,0:00:01.00,0:00:03.00,OP_JP,,0,0,0,,雨が上がっても',
    'Dialogue: 0,0:00:01.00,0:00:01.04,OP_JP,,0,0,0,,{\\pos(960,1068)\\clip(m 1 1)}雨が上がっても',
    'Dialogue: 0,0:00:01.04,0:00:01.08,OP_JP,,0,0,0,,{\\pos(960,1068)\\clip(m 2 2)}雨が上がっても',
    'Dialogue: 0,0:00:01.08,0:00:03.00,OP_JP,,0,0,0,,{\\pos(960,1068)\\clip(m 3 3)}雨が上がっても',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');

  assert.deepEqual(cues, [
    {
      startTime: 1,
      endTime: 3,
      text: '雨が上がっても',
      source: 'canonical-ass',
      animationStartTime: 1,
      animationEndTime: 3,
    },
  ]);
});

test('parseSubtitleCues collapses tag-less animation frames in converted SRT', () => {
  // ASS -> SRT conversion drops override tags, so only the ~0.04s frame timing remains.
  const lines = ['1', '00:00:07,870 --> 00:00:07,910', 'Kaguya Wants to be Confessed to', ''];
  for (let i = 1; i < 8; i++) {
    const start = 7910 + (i - 1) * 40;
    const end = start + 40;
    const at = (ms: number) =>
      `00:00:0${Math.floor(ms / 1000)},${String(ms % 1000).padStart(3, '0')}`;
    lines.push(String(i + 1), `${at(start)} --> ${at(end)}`, 'Kaguya Wants to be Confessed to', '');
  }

  const cues = parseSubtitleCues(lines.join('\n'), 'test.srt');

  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.startTime, 7.87);
});

test('parseSubtitleCues keeps identical lines that recur far apart', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:02.00,Default,,0,0,0,,なんで',
    'Dialogue: 0,0:05:00.00,0:05:01.00,Default,,0,0,0,,なんで',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');

  assert.equal(cues.length, 2);
  assert.equal(cues[0]!.startTime, 1.0);
  assert.equal(cues[1]!.startTime, 300.0);
});

test('parseSubtitleCues keeps two positioned signs that repeat the same text', () => {
  // Both carry override tags, but `\pos` and `\fad` are static placement, not animation.
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:01:00.00,0:01:03.00,Sign,,0,0,0,,{\\pos(960,120)\\fad(200,200)}第一話',
    'Dialogue: 0,0:01:03.00,0:01:06.00,Sign,,0,0,0,,{\\pos(960,900)\\fad(200,200)}第一話',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');

  assert.equal(cues.length, 2);
  assert.equal(cues[1]!.startTime, 63.0);
});

test('parseSubtitleCues keeps a run of ordinary positioned lines separate', () => {
  // Three events is a sequence, but none of them runs at animation-frame speed.
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:01:00.00,0:01:02.00,Sign,,0,0,0,,{\\pos(960,120)\\fad(100,100)}止まれ',
    'Dialogue: 0,0:01:02.00,0:01:04.00,Sign,,0,0,0,,{\\pos(960,120)\\fad(100,100)}止まれ',
    'Dialogue: 0,0:01:04.00,0:01:06.00,Sign,,0,0,0,,{\\pos(960,120)\\fad(100,100)}止まれ',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');

  assert.equal(cues.length, 3);
});

test('parseSubtitleCues keeps a short repeated SRT pair without burst evidence', () => {
  const content = [
    '1',
    '00:00:01,000 --> 00:00:01,200',
    'えっ',
    '',
    '2',
    '00:00:01,200 --> 00:00:01,400',
    'えっ',
    '',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.srt');

  assert.equal(cues.length, 2);
});

test('parseSubtitleCues collapses a burst marked only by the Effect column', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:01.05,OP_JP,,0,0,0,Karaoke,歌詞',
    'Dialogue: 0,0:00:01.05,0:00:01.09,OP_JP,,0,0,0,Karaoke,歌詞',
    'Dialogue: 0,0:00:01.09,0:00:03.55,OP_JP,,0,0,0,Karaoke,歌詞',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');

  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.endTime, 3.55);
});

test('parseSubtitleCues keeps a second karaoke burst that starts after a gap', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:01.05,OP_JP,,0,0,0,,{\\clip(m 1 1)}リフレイン',
    'Dialogue: 0,0:00:01.05,0:00:01.09,OP_JP,,0,0,0,,{\\clip(m 2 2)}リフレイン',
    'Dialogue: 0,0:00:01.09,0:00:03.00,OP_JP,,0,0,0,,{\\clip(m 3 3)}リフレイン',
    'Dialogue: 0,0:00:20.00,0:00:20.05,OP_JP,,0,0,0,,{\\clip(m 1 1)}リフレイン',
    'Dialogue: 0,0:00:20.05,0:00:20.09,OP_JP,,0,0,0,,{\\clip(m 2 2)}リフレイン',
    'Dialogue: 0,0:00:20.09,0:00:22.00,OP_JP,,0,0,0,,{\\clip(m 3 3)}リフレイン',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');

  assert.equal(cues.length, 2);
  assert.equal(cues[0]!.endTime, 3.0);
  assert.equal(cues[1]!.startTime, 20.0);
  assert.equal(cues[1]!.endTime, 22.0);
});

test('parseSubtitleCues does not merge a burst into unrelated dialogue between frames', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:01.05,OP_JP,,0,0,0,,{\\clip(m 1 1)}歌詞',
    'Dialogue: 0,0:00:01.02,0:00:03.00,Dial_JP,,0,0,0,,別のセリフ',
    'Dialogue: 0,0:00:01.05,0:00:01.09,OP_JP,,0,0,0,,{\\clip(m 2 2)}歌詞',
    'Dialogue: 0,0:00:01.09,0:00:03.55,OP_JP,,0,0,0,,{\\clip(m 3 3)}歌詞',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');

  assert.equal(cues.length, 2);
  assert.deepEqual(
    cues.map((cue) => cue.text),
    ['歌詞', '別のセリフ'],
  );
  assert.equal(cues[0]!.endTime, 3.55);
});

test('parseSubtitleCues keeps rapid ASS lines from different actors separate', () => {
  // Three 200ms `えっ` reactions traded between characters. Fast, adjacent and identical,
  // but authored as three lines: different styles and different actors.
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:01.20,Dial_A,アリス,0,0,0,,えっ',
    'Dialogue: 0,0:00:01.20,0:00:01.40,Dial_B,ボブ,0,0,0,,えっ',
    'Dialogue: 0,0:00:01.40,0:00:01.60,Dial_C,キャロル,0,0,0,,えっ',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');

  assert.equal(cues.length, 3);
});

test('parseSubtitleCues reads the speaker column when it is spelled Actor', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Actor, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:01.20,Dial_JP,アリス,0,0,0,,えっ',
    'Dialogue: 0,0:00:01.20,0:00:01.40,Dial_JP,ボブ,0,0,0,,えっ',
    'Dialogue: 0,0:00:01.40,0:00:01.60,Dial_JP,キャロル,0,0,0,,えっ',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');

  assert.equal(cues.length, 3);
});

test('parseSubtitleCues does not treat a custom Effect name as animation', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:01.20,Sign,,0,0,0,scrolling-credit,制作',
    'Dialogue: 0,0:00:01.20,0:00:01.40,Sign,,0,0,0,scrolling-credit,制作',
    'Dialogue: 0,0:00:01.40,0:00:01.60,Sign,,0,0,0,scrolling-credit,制作',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');

  assert.equal(cues.length, 3);
});

test('parseSubtitleCues keeps rapid ASS lines that share a style but not an actor', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:01.20,Dial_JP,アリス,0,0,0,,えっ',
    'Dialogue: 0,0:00:01.20,0:00:01.40,Dial_JP,ボブ,0,0,0,,えっ',
    'Dialogue: 0,0:00:01.40,0:00:01.60,Dial_JP,キャロル,0,0,0,,えっ',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');

  assert.equal(cues.length, 3);
});

test('parseSubtitleCues keeps untagged rapid ASS repeats separate', () => {
  // No overrides at all: timing-only evidence is an SRT/VTT fallback and must not apply
  // to ASS, where the absence of typesetting is itself evidence of plain dialogue.
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:01.05,Dial_JP,,0,0,0,,えっ',
    'Dialogue: 0,0:00:01.05,0:00:01.10,Dial_JP,,0,0,0,,えっ',
    'Dialogue: 0,0:00:01.10,0:00:01.15,Dial_JP,,0,0,0,,えっ',
    'Dialogue: 0,0:00:01.15,0:00:01.20,Dial_JP,,0,0,0,,えっ',
    'Dialogue: 0,0:00:01.20,0:00:01.25,Dial_JP,,0,0,0,,えっ',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');

  assert.equal(cues.length, 5);
});

test('parseSubtitleCues keeps repeated signs sharing one static clip', () => {
  // `\clip` is a static shape for the event. Three events with the identical clip were
  // typeset the same way, so none of them is a frame of the others.
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:01.20,Sign,,0,0,0,,{\\clip(0,0,100,100)}注意',
    'Dialogue: 0,0:00:01.20,0:00:01.40,Sign,,0,0,0,,{\\clip(0,0,100,100)}注意',
    'Dialogue: 0,0:00:01.40,0:00:01.60,Sign,,0,0,0,,{\\clip(0,0,100,100)}注意',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');

  assert.equal(cues.length, 3);
});

test('parseSubtitleCues collapses a sign animated through \\t', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:01.20,Sign,,0,0,0,,{\\pos(10,10)\\t(0,200,\\frz30)}回る',
    'Dialogue: 0,0:00:01.20,0:00:01.40,Sign,,0,0,0,,{\\pos(10,10)\\t(0,200,\\frz30)}回る',
    'Dialogue: 0,0:00:01.40,0:00:03.00,Sign,,0,0,0,,{\\pos(10,10)\\t(0,200,\\frz30)}回る',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');

  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.endTime, 3.0);
});

test('parseSubtitleCues keeps a short repeated SRT run above the frame threshold', () => {
  // Five contiguous 200ms cues: a sequence, but nowhere near animation-frame speed.
  const lines: string[] = [];
  for (let i = 0; i < 5; i++) {
    const start = 1000 + i * 200;
    const at = (ms: number) =>
      `00:00:0${Math.floor(ms / 1000)},${String(ms % 1000).padStart(3, '0')}`;
    lines.push(String(i + 1), `${at(start)} --> ${at(start + 200)}`, 'えっ', '');
  }

  const cues = parseSubtitleCues(lines.join('\n'), 'test.srt');

  assert.equal(cues.length, 5);
});

test('parseSubtitleCues keeps a short SRT frame run below the minimum length', () => {
  // Four 40ms frames: frame-speed, but too few to tell an animation from an artefact.
  const lines: string[] = [];
  for (let i = 0; i < 4; i++) {
    const start = 7870 + i * 40;
    const at = (ms: number) =>
      `00:00:0${Math.floor(ms / 1000)},${String(ms % 1000).padStart(3, '0')}`;
    lines.push(String(i + 1), `${at(start)} --> ${at(start + 40)}`, 'タイトル', '');
  }

  const cues = parseSubtitleCues(lines.join('\n'), 'test.srt');

  assert.equal(cues.length, 4);
});

test('parseSubtitleCues applies ASS burst rules to ASS content behind an .srt filename', () => {
  // The extension lies, so the SRT parser finds nothing and the content-sniffing fallback
  // takes over -- which has to carry the `ass` source format with it, or the far stricter
  // timing-only thresholds would let this karaoke burst through as three cues.
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:01.20,Karaoke,,0,0,0,,{\\k20}歌詞',
    'Dialogue: 0,0:00:01.20,0:00:01.40,Karaoke,,0,0,0,,{\\k20}歌詞',
    'Dialogue: 0,0:00:01.40,0:00:03.00,Karaoke,,0,0,0,,{\\k20}歌詞',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.srt');

  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.startTime, 1.0);
  assert.equal(cues[0]!.endTime, 3.0);
  assert.equal(cues[0]!.text, '歌詞');
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

test('parseSubtitleCues skips zero-duration ASS metadata events', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:00.00,0:00:00.00,Default,,0,0,0,,[Script Info]',
    'Dialogue: 0,0:00:01.00,0:00:02.00,Default,,0,0,0,,Real subtitle',
  ].join('\n');

  assert.deepEqual(parseSubtitleCues(content, 'test.ass'), [
    { startTime: 1, endTime: 2, text: 'Real subtitle' },
  ]);
});

test('parseSubtitleCues drops malformed ASS spacer reset debris', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:02.00,Background,,0,0,0,,{\\pos(10,10)}\\h\\h\\h\\{\\fr0',
    'Dialogue: 1,0:00:01.00,0:00:02.00,Default,,0,0,0,,Visible line',
  ].join('\n');

  assert.deepEqual(parseSubtitleCues(content, 'test.ass'), [
    { startTime: 1, endTime: 2, text: 'Visible line' },
  ]);
});

test('parseSubtitleCues recovers spaces encoded only by positioned Latin glyph gaps', () => {
  const glyphs = [
    ['T', 100],
    ['h', 118],
    ['e', 136],
    ['s', 164],
    ['t', 178],
    ['a', 194],
    ['r', 210],
    ['s', 227],
    ['I', 255],
    ['s', 275],
    ['e', 293],
    ['e', 311],
  ] as const;
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    ...[0, 1].flatMap((layer) =>
      glyphs.map(
        ([glyph, x], index) =>
          `Dialogue: ${layer},0:00:01.00,0:00:04.00,OP English,,0,0,0,,{\\pos(${x},110)\\t(${index * 2},${index * 2 + 100},\\fscx120)}${glyph}`,
      ),
    ),
  ].join('\n');

  assert.equal(parseSubtitleCues(content, 'test.ass')[0]?.text, 'The stars I see');
});

test('parseSubtitleCues adds a missing word space after positioned punctuation', () => {
  const fragments = [
    ['H', 100],
    ['i,', 119],
    ['t', 153],
    ['h', 168],
    ['e', 186],
    ['r', 202],
    ['e', 216],
  ] as const;
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    ...[0, 1].flatMap((layer) =>
      fragments.map(
        ([fragment, x], index) =>
          `Dialogue: ${layer},0:00:01.00,0:00:04.00,OP English,,0,0,0,,{\\pos(${x},110)\\t(${index * 2},${index * 2 + 100},\\fscx120)}${fragment}`,
      ),
    ),
  ].join('\n');

  assert.equal(parseSubtitleCues(content, 'test.ass')[0]?.text, 'Hi, there');
});

test('parseSubtitleCues does not split a positioned thousands separator', () => {
  const fragments = [
    ['1,', 100],
    ['000', 145],
    ['0', 185],
    ['0', 205],
    ['0', 225],
  ] as const;
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    ...[0, 1].flatMap((layer) =>
      fragments.map(
        ([fragment, x], index) =>
          `Dialogue: ${layer},0:00:01.00,0:00:04.00,OP English,,0,0,0,,{\\pos(${x},110)\\t(${index * 2},${index * 2 + 100},\\fscx120)}${fragment}`,
      ),
    ),
  ].join('\n');

  assert.equal(parseSubtitleCues(content, 'test.ass')[0]?.text, '1,000000');
});

test('parseSubtitleCues does not split a wide glyph from its punctuated suffix', () => {
  const fragments = [
    ['v', 904],
    ['o', 924],
    ['i', 939],
    ['c', 955],
    ['e', 976],
    ['r', 1004],
    ['e', 1021],
    ['a', 1042],
    ['c', 1063],
    ['h', 1083],
    ['e', 1104],
    ['d', 1125],
    ['m', 1161],
    ['e,', 1193],
  ] as const;
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    ...[0, 1].flatMap((layer) =>
      fragments.map(
        ([fragment, x], index) =>
          `Dialogue: ${layer},0:00:01.00,0:00:04.00,OP English,,0,0,0,,{\\pos(${x},110)\\t(${index * 2},${index * 2 + 100},\\fscx120)}${fragment}`,
      ),
    ),
  ].join('\n');

  assert.equal(parseSubtitleCues(content, 'test.ass')[0]?.text, 'voice reached me,');
});

test('parseSubtitleCues spaces positioned lyric fragments across authored rows', () => {
  const fragments = [
    ['My', 472, 39],
    ['song!', 543, 39],
    ['My', 507, 78],
    ['song!', 578, 78],
    ['ku', 643, 39],
    ['chi', 683, 39],
    ['zu', 722, 39],
    ['sa', 757, 39],
    ['n', 783, 39],
    ['de', 811, 39],
  ] as const;
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    ...[0, 1].flatMap((layer) =>
      fragments.map(
        ([fragment, x, y], index) =>
          `Dialogue: ${layer},0:00:01.00,0:00:04.00,OP Romaji,,0,0,0,,{\\pos(${x},${y})\\t(${index * 2},${index * 2 + 100},\\fscx120)}${fragment}`,
      ),
    ),
  ].join('\n');

  assert.equal(parseSubtitleCues(content, 'test.ass')[0]?.text, 'My song! My song! kuchizusande');
});

test('parseSubtitleCues recovers positioned word gaps between romaji fragments', () => {
  const fragments = [
    ['sa', 380],
    ['ga', 421],
    ['shi', 467],
    ['te', 510],
    ['ta', 545],
    ['ha', 593],
    ['ji', 624],
    ['ke', 655],
    ['ta', 693],
    ['i', 726],
    ['ro', 749],
    ['no', 798],
    ['yu', 849],
    ['me', 895],
  ] as const;
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    ...[0, 1].flatMap((layer) =>
      fragments.map(
        ([fragment, x], index) =>
          `Dialogue: ${layer},0:00:01.00,0:00:04.00,OP Romaji,,0,0,0,,{\\pos(${x},110)\\t(${index * 2},${index * 2 + 100},\\fscx120)}${fragment}`,
      ),
    ),
  ].join('\n');

  assert.equal(parseSubtitleCues(content, 'test.ass')[0]?.text, 'sagashiteta hajiketa iro no yume');
});

test('parseSubtitleCues recovers clear word gaps in a short romaji line', () => {
  const fragments = [
    ['bo', 542],
    ['ku', 584],
    ['wo', 640],
    ['yo', 697],
    ['bu', 738],
  ] as const;
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    ...[0, 1].flatMap((layer) =>
      fragments.map(
        ([fragment, x], index) =>
          `Dialogue: ${layer},0:00:01.00,0:00:04.00,OP Romaji,,0,0,0,,{\\pos(${x},110)\\t(${index * 2},${index * 2 + 100},\\fscx120)}${fragment}`,
      ),
    ),
  ].join('\n');

  assert.equal(parseSubtitleCues(content, 'test.ass')[0]?.text, 'boku wo yobu');
});

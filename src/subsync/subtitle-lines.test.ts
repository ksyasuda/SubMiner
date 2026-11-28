import test from 'node:test';
import assert from 'node:assert/strict';
import {
  applyTimingToSubtitleFile,
  decodeSubtitleBytes,
  decodeSubtitleBytesWithStatus,
  extractSpansFromParsedSubtitle,
  formatSpansAsSrt,
  isSignStyleOrEffect,
  isSignsOrSongsTrackMeta,
  parseSubtitleFileContent,
} from './subtitle-lines';

test('isSignStyleOrEffect separates signs/songs styles from dialogue styles', () => {
  for (const dialogueStyle of [
    'Default',
    'Subtitle',
    'Subtitles',
    'Notebook',
    'Main Top',
    'DefaultItalics',
    'Edward',
    'Ed',
    'Op',
    'TL',
    'Caption',
  ]) {
    assert.equal(isSignStyleOrEffect(dialogueStyle), false, `${dialogueStyle} should be dialogue`);
  }

  for (const signStyle of [
    'Sign',
    'Signs',
    'sign_0423_EpTitle',
    'Title',
    'Episode Title',
    'OP Kara',
    'ED-Romaji',
    'OPRomaji',
    'TLNote',
    'TL Note',
    'Insert Song',
    'Lyrics',
    'Typeset',
    'ED',
    '看板',
    'OP歌詞',
    'タイトル',
  ]) {
    assert.equal(isSignStyleOrEffect(signStyle), true, `${signStyle} should be a sign/song`);
  }
});

test('isSignsOrSongsTrackMeta flags signs, songs, lyrics, and forced tracks', () => {
  assert.equal(isSignsOrSongsTrackMeta({ forced: true, title: 'English' }), true);
  assert.equal(isSignsOrSongsTrackMeta({ title: 'English (Signs & Songs)' }), true);
  assert.equal(isSignsOrSongsTrackMeta({ title: 'S&S' }), true);
  assert.equal(isSignsOrSongsTrackMeta({ 'external-filename': '/tmp/ep01.forced.ass' }), true);
  assert.equal(isSignsOrSongsTrackMeta({ title: 'English [Full Dialogue]' }), false);
  assert.equal(isSignsOrSongsTrackMeta({ title: 'Japanese' }), false);
  for (const fullTitle of [
    'English (Full + Songs)',
    'Dialogue & Songs',
    'Full Subtitles [with lyrics]',
    '日本語 (歌詞あり)',
  ]) {
    assert.equal(isSignsOrSongsTrackMeta({ title: fullTitle }), false, fullTitle);
  }
});

test('parseSubtitleFileContent classifies ASS signs, drawings, karaoke, and preserves format on rewrite', () => {
  const ass = [
    '[Script Info]',
    'Title: Test Episode',
    '',
    '[V4+ Styles]',
    'Format: Name, Fontname',
    'Style: Default,Arial',
    '',
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Comment: 0,0:00:00.00,0:00:05.00,Default,,0,0,0,,keep this comment',
    'Dialogue: 0,0:00:10.00,0:00:12.50,Default,,0,0,0,,{\\an8}待って！',
    'Dialogue: 0,0:00:11.00,0:00:13.00,Sign,,0,0,0,,{\\pos(100,200)}木ノ葉病院',
    'Dialogue: 0,0:00:13.00,0:00:15.00,Default,,0,0,0,,{\\pos(320,40)}看板',
    'Dialogue: 0,0:00:16.00,0:00:18.00,OPRomaji,,0,0,0,,{\\k20}ha{\\k30}ru',
    'Dialogue: 0,0:00:19.00,0:00:21.00,Default,,0,0,0,,{\\p1}m 0 0 l 100 0{\\p0}',
    'Dialogue: 0,0:00:22.00,0:00:24.00,Default,,0,0,0,,そうか, わかった',
  ].join('\r\n');

  const parsed = parseSubtitleFileContent(ass, 'ep01.ass');
  assert.equal(parsed.format, 'ass');
  assert.equal(parsed.eol, '\r\n');
  assert.equal(parsed.events.length, 6);

  const { spans, dialogueIndices, dialogueCount, signCount } = extractSpansFromParsedSubtitle(
    parsed,
    { stripSigns: true },
  );
  assert.equal(dialogueCount, 2);
  assert.equal(signCount, 4);
  assert.deepEqual(dialogueIndices, [0, 5]);
  assert.deepEqual(spans, [
    { start: 10, end: 12.5 },
    { start: 22, end: 24 },
  ]);

  // Apply different shifts to the two dialogue lines (+2s and +5s) and verify signs follow nearest dialogue.
  const rewritten = applyTimingToSubtitleFile(parsed, dialogueIndices, [2.0, 5.0], 1.0);
  assert.ok(
    rewritten.includes('Comment: 0,0:00:00.00,0:00:05.00,Default,,0,0,0,,keep this comment'),
  );
  assert.ok(
    rewritten.includes('Dialogue: 0,0:00:12.00,0:00:14.50,Default,,0,0,0,,{\\an8}待って！'),
  );
  // Sign at 11s is nearest to dialogue 0 (10s) -> +2s => 13s..15s
  assert.ok(
    rewritten.includes('Dialogue: 0,0:00:13.00,0:00:15.00,Sign,,0,0,0,,{\\pos(100,200)}木ノ葉病院'),
  );
  // Drawing at 19s is nearest to dialogue 5 (22s) -> +5s => 24s..26s
  assert.ok(
    rewritten.includes(
      'Dialogue: 0,0:00:24.00,0:00:26.00,Default,,0,0,0,,{\\p1}m 0 0 l 100 0{\\p0}',
    ),
  );
  // Commas inside Text field stay intact!
  assert.ok(
    rewritten.includes('Dialogue: 0,0:00:27.00,0:00:29.00,Default,,0,0,0,,そうか, わかった'),
  );
  assert.ok(rewritten.includes('\r\n'));
});

test('parseSubtitleFileContent keeps positioned ASS lines as dialogue when most lines use \\pos', () => {
  const ass = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:03.00,Default,,0,0,0,,{\\pos(640,680)}もしもし',
    'Dialogue: 0,0:00:04.00,0:00:06.00,Default,,0,0,0,,{\\pos(640,40)}元気？',
  ].join('\n');

  const parsed = parseSubtitleFileContent(ass, 'all-pos.ass');
  const extracted = extractSpansFromParsedSubtitle(parsed, { stripSigns: true });
  assert.equal(extracted.dialogueCount, 2);
  assert.equal(extracted.signCount, 0);
});

test('parseSubtitleFileContent preserves SRT/VTT cue settings and strips music-only lines', () => {
  const vtt = [
    'WEBVTT',
    '',
    '00:01.000 --> 00:02.500 align:center line:90%',
    'こんにちは',
    '',
    '00:03.000 --> 00:05.000',
    '♪ La la la ♪',
    '',
    '00:06.000 --> 00:08.000',
    '元気ですか？',
  ].join('\n');

  const parsed = parseSubtitleFileContent(vtt, 'ep01.vtt');
  assert.equal(parsed.format, 'vtt');
  const extracted = extractSpansFromParsedSubtitle(parsed, { stripSigns: true });
  assert.equal(extracted.dialogueCount, 2);
  assert.equal(extracted.signCount, 1);

  const rewritten = applyTimingToSubtitleFile(parsed, extracted.dialogueIndices, [1.5, 1.5], 1.0);
  assert.ok(rewritten.includes('WEBVTT'));
  assert.ok(rewritten.includes('00:00:02.500 --> 00:00:04.000 align:center line:90%'));
  assert.ok(rewritten.includes('00:00:04.500 --> 00:00:06.500\n♪ La la la ♪'));
});

test('decodeSubtitleBytesWithStatus marks lossless UTF-8 and lossy undecodable bytes accurately', () => {
  const bomBytes = new Uint8Array([0xef, 0xbb, 0xbf, ...Buffer.from('hello', 'utf8')]);
  assert.equal(decodeSubtitleBytes(bomBytes), 'hello');
  assert.deepEqual(decodeSubtitleBytesWithStatus(bomBytes), {
    text: '\uFEFFhello',
    lossless: true,
  });
  assert.deepEqual(decodeSubtitleBytesWithStatus(Buffer.from('hello', 'utf8')), {
    text: 'hello',
    lossless: true,
  });

  // Bytes invalid in UTF-8, Shift_JIS, and EUC-JP trigger lossy fallback.
  const corrupt = new Uint8Array([0x80, 0xff, 0xfe, 0x80]);
  const lossy = decodeSubtitleBytesWithStatus(corrupt);
  assert.equal(lossy.lossless, false);

  // Ambiguous bytes valid in both Shift_JIS and EUC-JP (e.g. EUC-JP 'あ' = A4 A2) must not be marked lossless.
  const ambiguousEucJp = new Uint8Array([0xa4, 0xa2]);
  const ambiguousResult = decodeSubtitleBytesWithStatus(ambiguousEucJp);
  assert.equal(ambiguousResult.lossless, false);

  // Unambiguous Shift_JIS bytes (0x82 0xA0 = 'あ', invalid in UTF-8 and EUC-JP) remain lossless.
  const unambiguousSjis = new Uint8Array([0x82, 0xa0]);
  assert.deepEqual(decodeSubtitleBytesWithStatus(unambiguousSjis), {
    text: '\uFEFFあ',
    lossless: true,
  });

  const srt = formatSpansAsSrt([
    { start: 5.5, end: 7.0 },
    { start: 1.25, end: 2.75 },
  ]);
  assert.ok(srt.includes('00:00:01,250 --> 00:00:02,750'));
  assert.ok(srt.includes('00:00:05,500 --> 00:00:07,000'));
});

test('applyTimingToSubtitleFile writes a BOM only when the source had one or was not UTF-8', () => {
  const srt = '1\n00:00:01,000 --> 00:00:02,000\nこんにちは\n';
  const rewrite = (bytes: Uint8Array): string => {
    const parsed = parseSubtitleFileContent(decodeSubtitleBytesWithStatus(bytes).text, 'a.srt');
    return applyTimingToSubtitleFile(parsed, [0], [1], 1.0);
  };

  const plain = rewrite(Buffer.from(srt, 'utf8'));
  assert.ok(plain.startsWith('1\n00:00:02,000 --> 00:00:03,000'));
  assert.ok(
    rewrite(new Uint8Array([0xef, 0xbb, 0xbf, ...Buffer.from(srt, 'utf8')])).startsWith(
      '\uFEFF1\n',
    ),
  );
  // 'あ' in Shift_JIS (invalid UTF-8) is rewritten as UTF-8 with a BOM.
  const sjis = new Uint8Array([
    ...Buffer.from('1\n00:00:01,000 --> 00:00:02,000\n', 'ascii'),
    0x82,
    0xa0,
  ]);
  assert.ok(rewrite(sjis).startsWith('\uFEFF1\n00:00:02,000'));
});

test('signs and songs follow the dialogue line nearest in time, not in file order', () => {
  const ass = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:20.00,0:00:22.00,Default,,0,0,0,,一',
    'Dialogue: 0,0:00:40.00,0:00:42.00,Default,,0,0,0,,二',
    'Dialogue: 0,0:10:00.00,0:10:02.00,Default,,0,0,0,,三',
    'Dialogue: 0,0:10:20.00,0:10:22.00,Default,,0,0,0,,四',
    // Typesetting kept at the bottom of the file, timed near the first act.
    'Dialogue: 0,0:00:10.00,0:00:12.00,OP Kara,,0,0,0,,{\\k20}ら',
    'Dialogue: 0,0:01:30.00,0:01:32.00,Sign,,0,0,0,,看板',
  ].join('\n');
  const parsed = parseSubtitleFileContent(ass, 'ep.ass');
  const { dialogueIndices } = extractSpansFromParsedSubtitle(parsed);
  assert.deepEqual(dialogueIndices, [0, 1, 2, 3]);

  const rewritten = applyTimingToSubtitleFile(parsed, dialogueIndices, [0, 0, 5, 5], 1.0);
  assert.ok(rewritten.includes('0:00:10.00,0:00:12.00,OP Kara'));
  assert.ok(rewritten.includes('0:01:30.00,0:01:32.00,Sign'));
  assert.ok(rewritten.includes('0:10:05.00,0:10:07.00,Default'));
});

test('only lines that open with a music mark count as lyrics', () => {
  const srt = [
    '1',
    '00:00:01,000 --> 00:00:02,000',
    'ありがとう♪',
    '',
    '2',
    '00:00:03,000 --> 00:00:04,000',
    '♪ 夢の中へ ♪',
    '',
  ].join('\n');
  const parsed = parseSubtitleFileContent(srt, 'a.srt');
  assert.deepEqual(
    parsed.events.map((event) => event.sign),
    [false, true],
  );
});

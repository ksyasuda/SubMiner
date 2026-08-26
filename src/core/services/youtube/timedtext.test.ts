import assert from 'node:assert/strict';
import test from 'node:test';
import { convertYoutubeTimedTextToVtt, normalizeYoutubeAutoVtt } from './timedtext';

test('convertYoutubeTimedTextToVtt leaves malformed numeric entities literal', () => {
  const result = convertYoutubeTimedTextToVtt(
    '<timedtext><body><p t="0" d="1000">&#99999999; &#x110000; &#x41;</p></body></timedtext>',
  );

  assert.equal(
    result,
    ['WEBVTT', '', '00:00:00.000 --> 00:00:01.000', '&#99999999; &#x110000; A', ''].join('\n'),
  );
});

test('convertYoutubeTimedTextToVtt does not swallow text after zero-length overlap rows', () => {
  const result = convertYoutubeTimedTextToVtt(
    [
      '<timedtext><body>',
      '<p t="0" d="2000">今日は</p>',
      '<p t="1000" d="0">今日はいい天気ですね</p>',
      '<p t="1000" d="2000">今日はいい天気ですね</p>',
      '</body></timedtext>',
    ].join(''),
  );

  assert.equal(
    result,
    [
      'WEBVTT',
      '',
      '00:00:00.000 --> 00:00:00.999',
      '今日は',
      '',
      '00:00:01.000 --> 00:00:03.000',
      'いい天気ですね',
      '',
    ].join('\n'),
  );
});

test('convertYoutubeTimedTextToVtt extends rolling captions to the next window event', () => {
  // Real-world shape of YouTube's sentence-level auto captions: window-append
  // filler rows (a="1", sometimes without d) mark the display timeline, while
  // long text rows carry a placeholder d="3000" far shorter than the speech.
  const result = convertYoutubeTimedTextToVtt(
    [
      '<timedtext><body>',
      '<p t="98550" d="3010" w="1" a="1">\n</p>',
      '<p t="98560" d="3000" w="1"><s ac="0">ありがとうって言えないよね。こんなんじゃ。</s></p>',
      '<p t="106950" w="1" a="1">\n</p>',
      '<p t="106960" d="3799" w="1"><s ac="0">私だったら無理だよ。</s></p>',
      '</body></timedtext>',
    ].join('\n'),
  );

  assert.equal(
    result,
    [
      'WEBVTT',
      '',
      '00:01:38.560 --> 00:01:46.950',
      'ありがとうって言えないよね。こんなんじゃ。',
      '',
      '00:01:46.960 --> 00:01:50.759',
      '私だったら無理だよ。',
      '',
    ].join('\n'),
  );
});

test('normalizeYoutubeAutoVtt strips cumulative rolling-caption prefixes', () => {
  const result = normalizeYoutubeAutoVtt(
    [
      'WEBVTT',
      '',
      '00:00:01.000 --> 00:00:02.000',
      '今日は',
      '',
      '00:00:02.000 --> 00:00:03.000',
      '今日はいい天気ですね',
      '',
      '00:00:03.000 --> 00:00:04.000',
      '今日はいい天気ですね本当に',
      '',
    ].join('\n'),
  );

  assert.equal(
    result,
    [
      'WEBVTT',
      '',
      '00:00:01.000 --> 00:00:02.000',
      '今日は',
      '',
      '00:00:02.000 --> 00:00:03.000',
      'いい天気ですね',
      '',
      '00:00:03.000 --> 00:00:04.000',
      '本当に',
      '',
    ].join('\n'),
  );
});

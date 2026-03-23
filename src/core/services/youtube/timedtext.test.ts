import assert from 'node:assert/strict';
import test from 'node:test';
import { convertYoutubeTimedTextToVtt } from './timedtext';

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

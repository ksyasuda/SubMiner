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

test('convertYoutubeTimedTextToVtt pages oversized two-row rolling captions', () => {
  const text =
    'あの西に結構こう山田がスーパーアプローチしてるんだけど西気づかないからちょっとこっちも気づかない感じでこう接してあげようかなて思ってんだけどあの唇巻き込んじゃうしあの思ってることも全部縁に出ちゃって自分であちゃったって言っちゃうタイプなんで結構なんかこうドライなんだけどそこがおもろいよねみたいな';
  const result = convertYoutubeTimedTextToVtt(
    [
      '<timedtext format="3">',
      '<head>',
      '<ws id="1" mh="2" ju="0" sd="3"/>',
      '<wp id="1" ap="6" ah="20" av="100" rc="2" cc="40"/>',
      '</head>',
      '<body>',
      '<w t="0" id="1" wp="1" ws="1"/>',
      `<p t="60440" d="3000" w="1"><s ac="0">${text}</s></p>`,
      '<p t="72695" w="1" a="1">\n</p>',
      '</body>',
      '</timedtext>',
    ].join('\n'),
  );

  const cues = result
    .trim()
    .split(/\n\n/)
    .filter((block) => block.includes('-->'));
  const cueText = cues.map((cue) => cue.split('\n').slice(1).join('\n'));

  assert.equal(cues.length, 2);
  assert.deepEqual(
    cues.map((cue) => cue.split('\n')[0]),
    ['00:01:00.440 --> 00:01:07.064', '00:01:07.064 --> 00:01:12.695'],
  );
  assert.ok(cueText.every((page) => [...page].length <= 80));
  assert.equal(cueText.join(''), text);
});

test('convertYoutubeTimedTextToVtt leaves pop-on captions intact', () => {
  const result = convertYoutubeTimedTextToVtt(
    [
      '<timedtext format="3">',
      '<head>',
      '<ws id="1" mh="0"/>',
      '<wp id="1" rc="2" cc="4"/>',
      '</head>',
      '<body>',
      '<w t="0" id="1" wp="1" ws="1"/>',
      '<p t="1000" d="3000" w="1">abcdefghijklmnopqrst</p>',
      '</body>',
      '</timedtext>',
    ].join('\n'),
  );

  assert.equal(
    result,
    ['WEBVTT', '', '00:00:01.000 --> 00:00:04.000', 'abcdefghijklmnopqrst', ''].join('\n'),
  );
});

test('convertYoutubeTimedTextToVtt keeps explicit sound-cue durations in rolling documents', () => {
  const result = convertYoutubeTimedTextToVtt(
    [
      '<timedtext><body>',
      '<p t="20305" d="2020" w="1">[音楽]</p>',
      '<p t="26269" w="1" a="1">\n</p>',
      '<p t="26279" d="3000" w="1"><s ac="0">じゃあ、君からお願いします。</s></p>',
      '</body></timedtext>',
    ].join('\n'),
  );

  assert.equal(
    result,
    [
      'WEBVTT',
      '',
      '00:00:20.305 --> 00:00:22.325',
      '[音楽]',
      '',
      '00:00:26.279 --> 00:00:29.279',
      'じゃあ、君からお願いします。',
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

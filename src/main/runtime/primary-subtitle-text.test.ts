import assert from 'node:assert/strict';
import test from 'node:test';
import { parseSubtitleCues } from '../../core/services/subtitle-cue-parser';
import {
  resolveCanonicalPrimarySubtitle,
  resolvePrimarySubtitleText,
  stripCanonicalFragmentLines,
} from './primary-subtitle-text';

test('resolvePrimarySubtitleText collapses full-span ASS style layers through parsed cues', () => {
  const ass = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 3,0:22:40.05,0:22:45.76,EDJP,,0,0,0,,{\\fad(400,400)\\bord0\\blur0.8}鏡の奥まで目を凝らして',
    'Dialogue: 2,0:22:40.05,0:22:45.76,EDJP,,0,0,0,,{\\fad(400,400)}鏡の奥まで目を凝らして',
    'Dialogue: 1,0:22:40.05,0:22:45.76,EDJP,,0,0,0,,{\\fad(400,400)\\bord6}鏡の奥まで目を凝らして',
    'Dialogue: 0,0:22:40.05,0:22:45.76,EDJP,,0,0,0,,{\\fad(400,400)\\bord8\\blur4}鏡の奥まで目を凝らして',
  ].join('\n');
  const cues = parseSubtitleCues(ass, 'polar-opposites-s01e08.ass');

  assert.deepEqual(cues, [
    { startTime: 22 * 60 + 40.05, endTime: 22 * 60 + 45.76, text: '鏡の奥まで目を凝らして' },
  ]);

  assert.equal(
    resolvePrimarySubtitleText({
      liveText: [
        '鏡の奥まで目を凝らして',
        '鏡の奥まで目を凝らして',
        '鏡の奥まで目を凝らして',
        '鏡の奥まで目を凝らして',
      ].join('\n'),
      currentTimeSec: 22 * 60 + 44,
      cues,
    }),
    '鏡の奥まで目を凝らして',
  );
});

test('resolvePrimarySubtitleText keeps live text when active parsed cues do not explain it all', () => {
  const liveText = '普通のセリフ\n鏡の奥まで目を凝らして\n鏡の奥まで目を凝らして';

  assert.equal(
    resolvePrimarySubtitleText({
      liveText,
      currentTimeSec: 2,
      cues: [{ startTime: 1, endTime: 3, text: '鏡の奥まで目を凝らして' }],
    }),
    liveText,
  );
});

test('resolvePrimarySubtitleText combines unique simultaneous parsed cues', () => {
  assert.equal(
    resolvePrimarySubtitleText({
      liveText: '一行目\n一行目\n二行目\n二行目',
      currentTimeSec: 2,
      cues: [
        { startTime: 1, endTime: 3, text: '一行目' },
        { startTime: 1, endTime: 3, text: '二行目' },
      ],
    }),
    '一行目\n二行目',
  );
});

test('resolvePrimarySubtitleText collapses whitespace variants of one ASS lyric', () => {
  const ass = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 2,0:00:01.00,0:00:03.00,EDJP,,0,0,0,,少しだけ好きになる',
    'Dialogue: 1,0:00:01.00,0:00:03.00,EDJP,,0,0,0,,少しだけ\\h好きになる',
    'Dialogue: 0,0:00:01.00,0:00:03.00,EDJP,,0,0,0,,少しだけ　好きになる',
  ].join('\n');
  const cues = parseSubtitleCues(ass, 'polar-opposites-s01e10.ass');

  assert.deepEqual(
    cues.map((cue) => cue.text),
    ['少しだけ好きになる', '少しだけ 好きになる', '少しだけ　好きになる'],
  );
  assert.equal(
    resolvePrimarySubtitleText({
      liveText: ['少しだけ好きになる', '少しだけ 好きになる', '少しだけ　好きになる'].join('\n'),
      currentTimeSec: 2,
      cues,
    }),
    '少しだけ好きになる',
  );
});

test('resolvePrimarySubtitleText tolerates stale time-pos at a parsed cue edge', () => {
  assert.equal(
    resolvePrimarySubtitleText({
      liveText: '新しい行\n新しい行',
      currentTimeSec: 0.8,
      cues: [{ startTime: 1, endTime: 3, text: '新しい行' }],
    }),
    '新しい行',
  );
});

test('resolvePrimarySubtitleText prefers an active canonical cue over flattened mpv glyphs', () => {
  // mpv renders each simultaneously active ASS event on its own sub-text line.
  const text = resolvePrimarySubtitleText({
    liveText: '今\n今\n今\n手\n手\n手\nにある\nにある\nにある',
    currentTimeSec: 2,
    cues: [
      {
        startTime: 1.2,
        endTime: 3.8,
        text: '今　手にある',
        source: 'canonical-ass',
      },
    ],
  });

  assert.equal(text, '今　手にある');
});

test('resolvePrimarySubtitleText preserves live text outside canonical cue timing', () => {
  const text = resolvePrimarySubtitleText({
    liveText: '通常の会話',
    currentTimeSec: 8,
    cues: [
      {
        startTime: 1.2,
        endTime: 3.8,
        text: '今　手にある',
        source: 'canonical-ass',
      },
    ],
  });

  assert.equal(text, '通常の会話');
});

test('resolvePrimarySubtitleText keeps concurrent dialogue that is not part of the animation', () => {
  // An insert song's canonical window can overlap real dialogue on the same track.
  const text = resolvePrimarySubtitleText({
    liveText: '普通のセリフ\n今\n手にある',
    currentTimeSec: 2,
    cues: [
      {
        startTime: 1.2,
        endTime: 3.8,
        text: '今　手にある',
        source: 'canonical-ass',
      },
    ],
  });

  assert.equal(text, '普通のセリフ\n今\n手にある');
});

test('resolvePrimarySubtitleText keeps a fresh line starting just after the animation ended', () => {
  const text = resolvePrimarySubtitleText({
    liveText: '次のセリフ',
    currentTimeSec: 4.1,
    cues: [
      {
        startTime: 1.2,
        endTime: 3.8,
        text: '今　手にある',
        source: 'canonical-ass',
      },
    ],
  });

  assert.equal(text, '次のセリフ');
});

test('resolvePrimarySubtitleText survives overlapping frames of consecutive karaoke lines', () => {
  // Near a line boundary the previous line's exit frames and the next line's entrance
  // frames render together; neither line alone explains every live segment.
  const cues = [
    { startTime: 1.2, endTime: 3.8, text: '今　手にある', source: 'canonical-ass' as const },
    { startTime: 3.8, endTime: 6.4, text: '物差しでは', source: 'canonical-ass' as const },
  ];

  assert.equal(
    resolvePrimarySubtitleText({
      liveText: '手にある\n物差し\nでは',
      currentTimeSec: 3.6,
      cues,
    }),
    '今　手にある',
  );
  assert.equal(
    resolvePrimarySubtitleText({
      liveText: '手にある\n物差し\nでは',
      currentTimeSec: 3.9,
      cues,
    }),
    '物差しでは',
  );
});

test('resolvePrimarySubtitleText combines simultaneous canonical cues in source order', () => {
  const text = resolvePrimarySubtitleText({
    liveText: 'fir\nst\nsecond',
    currentTimeSec: 2,
    cues: [
      { startTime: 1, endTime: 3, text: 'first', source: 'canonical-ass' },
      { startTime: 1.5, endTime: 2.5, text: 'second', source: 'canonical-ass' },
    ],
  });

  assert.equal(text, 'first\nsecond');
});

test('resolvePrimarySubtitleText collapses whitespace variants of a canonical lyric', () => {
  assert.equal(
    resolvePrimarySubtitleText({
      liveText: '少しだけ好きになる\n少しだけ　好きになる',
      currentTimeSec: 2,
      cues: [
        { startTime: 1, endTime: 3, text: '少しだけ好きになる', source: 'canonical-ass' },
        { startTime: 1, endTime: 3, text: '少しだけ　好きになる', source: 'canonical-ass' },
      ],
    }),
    '少しだけ好きになる',
  );
});

test('resolveCanonicalPrimarySubtitle covers a nearby generated animation edge', () => {
  const cue = {
    startTime: 1.2,
    endTime: 3.8,
    text: '今　手にある',
    source: 'canonical-ass' as const,
  };
  const resolved = resolveCanonicalPrimarySubtitle({
    liveText: '今\n手にある',
    currentTimeSec: 0.8,
    cues: [cue],
  });

  assert.deepEqual(resolved, {
    text: '今　手にある',
    startTime: 1.2,
    endTime: 3.8,
    cues: [cue],
  });
});

test('resolveCanonicalPrimarySubtitle covers exit frames that outlive the authored timing', () => {
  // Real generated animations keep exit fragments on screen well past the authored
  // comment window; the recorded animation envelope is what makes them resolvable.
  const cue = {
    startTime: 1.2,
    endTime: 3.8,
    text: '今　手にある',
    source: 'canonical-ass' as const,
    animationStartTime: 0.8,
    animationEndTime: 5.6,
  };
  const resolved = resolveCanonicalPrimarySubtitle({
    liveText: '今\n手にある',
    currentTimeSec: 5.4,
    cues: [cue],
  });

  assert.deepEqual(resolved, {
    text: '今　手にある',
    startTime: 1.2,
    endTime: 3.8,
    cues: [cue],
  });
});

test('resolvePrimarySubtitleText handles late exit frames overlapping the next active line', () => {
  // The previous line's exit fragments can persist more than a second into the next
  // authored line. The next line supplies the text; the previous line's envelope
  // explains its lingering fragments.
  const cues = [
    {
      startTime: 1.2,
      endTime: 3.8,
      text: '今　手にある',
      source: 'canonical-ass' as const,
      animationStartTime: 0.8,
      animationEndTime: 5.6,
    },
    {
      startTime: 3.8,
      endTime: 6.4,
      text: '物差しでは',
      source: 'canonical-ass' as const,
      animationStartTime: 3.4,
      animationEndTime: 7.0,
    },
  ];

  assert.equal(
    resolvePrimarySubtitleText({
      liveText: '手にある\n手にある\n物差し\nでは',
      currentTimeSec: 5.2,
      cues,
    }),
    '物差しでは',
  );
});

test('resolveCanonicalPrimarySubtitle rejects unrelated live text at the animation edge', () => {
  const resolved = resolveCanonicalPrimarySubtitle({
    liveText: '次のセリフ',
    currentTimeSec: 4.1,
    cues: [
      {
        startTime: 1.2,
        endTime: 3.8,
        text: '今　手にある',
        source: 'canonical-ass',
      },
    ],
  });

  assert.equal(resolved, null);
});

test('stripCanonicalFragmentLines drops fragment lines but keeps concurrent dialogue', () => {
  const cues = [
    {
      startTime: 1.2,
      endTime: 3.8,
      text: '今　手にある',
      source: 'canonical-ass' as const,
    },
  ];

  assert.equal(
    stripCanonicalFragmentLines({
      liveText: '普通のセリフ\n今\n手にある',
      currentTimeSec: 2,
      cues,
    }),
    '普通のセリフ',
  );
  // No canonical cue nearby: nothing to strip.
  assert.equal(
    stripCanonicalFragmentLines({ liveText: '普通のセリフ\n今', currentTimeSec: 30, cues }),
    '普通のセリフ\n今',
  );
  // Everything matched (defensive): return the input rather than empty text.
  assert.equal(
    stripCanonicalFragmentLines({ liveText: '今\n手にある', currentTimeSec: 2, cues }),
    '今\n手にある',
  );
});

test('resolveCanonicalPrimarySubtitle picks the cue its fragments spell, not the nearest', () => {
  // In the gap between two authored spans, the next line sits closer in time while only
  // the previous line's exit fragments are on screen: the fragments decide.
  const cues = [
    {
      startTime: 1,
      endTime: 3,
      text: '今　手にある',
      source: 'canonical-ass' as const,
      animationStartTime: 0.6,
      animationEndTime: 3.9,
    },
    {
      startTime: 4,
      endTime: 6,
      text: '物差しでは',
      source: 'canonical-ass' as const,
      animationStartTime: 3.5,
      animationEndTime: 6.4,
    },
  ];

  assert.equal(
    resolveCanonicalPrimarySubtitle({ liveText: '手にある', currentTimeSec: 3.8, cues })?.text,
    '今　手にある',
  );
  // Fragments of both lines in the gap: both envelopes cover the moment (distance 0),
  // and the earlier line wins the tie while it is still animating out.
  assert.equal(
    resolveCanonicalPrimarySubtitle({ liveText: '手にある\n物差し', currentTimeSec: 3.8, cues })
      ?.text,
    '今　手にある',
  );
});

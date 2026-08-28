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
    '一行目\n\n二行目',
  );
});

test('resolvePrimarySubtitleText removes duplicate lines across multiline parsed cues', () => {
  assert.equal(
    resolvePrimarySubtitleText({
      liveText: 'First line\nSecond line\nFirst line',
      currentTimeSec: 2,
      cues: [
        { startTime: 1, endTime: 3, text: 'First line\nSecond line' },
        { startTime: 1, endTime: 3, text: 'First line' },
      ],
    }),
    'First line\nSecond line',
  );
});

test('resolvePrimarySubtitleText removes equivalent full-width duplicate lines', () => {
  assert.equal(
    resolvePrimarySubtitleText({
      liveText: '２０分５３秒\n20分53秒',
      currentTimeSec: 2,
      cues: [
        { startTime: 1, endTime: 3, text: '２０分５３秒' },
        { startTime: 1, endTime: 3, text: '20分53秒' },
      ],
    }),
    '２０分５３秒',
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

test('resolvePrimarySubtitleText combines parsed dialogue with a reconstructed lyric', () => {
  const text = resolvePrimarySubtitleText({
    liveText: '普通のセリフ\n今\n今\n手\n手\nにある\nにある',
    currentTimeSec: 2,
    cues: [
      { startTime: 1, endTime: 3, text: '普通のセリフ' },
      {
        startTime: 1.2,
        endTime: 3.8,
        text: '今　手にある',
        source: 'reconstructed-ass',
      },
    ],
  });

  assert.equal(text, '普通のセリフ\n\n今　手にある');
});

test('resolvePrimarySubtitleText uses fragment grids only to account for live sign pieces', () => {
  const text = resolvePrimarySubtitleText({
    liveText: 'Ordinary dialogue\nMaid\nCafe',
    currentTimeSec: 2,
    cues: [
      { startTime: 1, endTime: 3, text: 'Ordinary dialogue' },
      {
        startTime: 1,
        endTime: 3,
        text: 'MaidCafeMaidCafe',
        source: 'reconstructed-ass',
        assLayout: { kind: 'fragment-grid', sourceOrder: 2 },
      },
    ],
  });

  assert.equal(text, 'Ordinary dialogue');
});

test('resolvePrimarySubtitleText drops malformed ASS control debris from live text', () => {
  const cues = parseSubtitleCues(
    [
      '[Events]',
      'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
      'Dialogue: 0,0:00:01.00,0:00:03.00,Default,,0,0,0,,Visible line',
    ].join('\n'),
    'test.ass',
  );

  assert.equal(
    resolvePrimarySubtitleText({
      liveText: 'Visible line\n\\\n{\\fr0',
      currentTimeSec: 2,
      cues,
    }),
    'Visible line',
  );
});

test('resolvePrimarySubtitleText preserves SRT text that resembles ASS control debris', () => {
  const liveText = 'Visible line\n\\\n{\\fr0';
  const cues = parseSubtitleCues(
    ['1', '00:00:01,000 --> 00:00:03,000', liveText].join('\n'),
    'test.srt',
  );

  assert.equal(resolvePrimarySubtitleText({ liveText, currentTimeSec: 2, cues }), liveText);
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

  assert.equal(text, 'first\n\nsecond');
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

test('resolvePrimarySubtitleText suppresses a live glyph wall when no cues are available', () => {
  const wall = [...'wansdumretoikhI'].join('\n');
  assert.equal(
    resolvePrimarySubtitleText({ liveText: `${wall}\ntai`, currentTimeSec: 1355, cues: null }),
    '',
  );
});

test('stripCanonicalFragmentLines drops a live glyph wall with no nearby canonical cues', () => {
  const wall = [...'wansdumretoikhI'].join('\n');
  assert.equal(
    stripCanonicalFragmentLines({
      liveText: `${wall}\nそれよりも　ノート…`,
      currentTimeSec: 1355,
      cues: [],
    }),
    'それよりも　ノート…',
  );
});

test('resolvePrimarySubtitleText keeps a line joining an active cue despite stale time-pos', () => {
  // Issue #220: mpv publishes the combined sub-text the moment a joining line's first
  // frame renders, while the observed time-pos still sits just before that line's
  // start. The joining cue must not be filtered out as inactive.
  assert.equal(
    resolvePrimarySubtitleText({
      liveText: 'Балда! Балда, балда, балда!\nСестренка не может остановиться',
      currentTimeSec: 767.78,
      cues: [
        { startTime: 767.19, endTime: 772.78, text: 'Балда! Балда, балда, балда!' },
        { startTime: 767.79, endTime: 771.15, text: 'Сестренка не может остановиться' },
      ],
    }),
    'Балда! Балда, балда, балда!\n\nСестренка не может остановиться',
  );
});

test('resolvePrimarySubtitleText drops a finished lyric whose exit ghosts outlive it beside a raw line', () => {
  // The reconstructed lyric ended at 6.0 but its exit ghost glyphs stay in the live
  // text until 7.0, while the next authored line is a plain raw event. The retired cue
  // must explain the ghost fragments without re-surfacing next to the active line.
  const cues = [
    {
      startTime: 1.0,
      endTime: 6.0,
      text: 'エネルギーはサイクル',
      source: 'reconstructed-ass' as const,
      animationStartTime: 0.5,
      animationEndTime: 7.0,
      assStyle: 'OP - JP',
    },
    { startTime: 6.0, endTime: 12.0, text: '象徴的なパレード' },
  ];

  assert.equal(
    resolvePrimarySubtitleText({
      liveText: 'エ\nネ\nル\nギ\nー\n象徴的なパレード',
      currentTimeSec: 6.5,
      cues,
    }),
    '象徴的なパレード',
  );
});

test('resolvePrimarySubtitleText stacks simultaneous cues by screen position, not start order', () => {
  // A top-anchored lyric and bottom dialogue: mpv draws the lyric above the dialogue for
  // the whole overlap. Whichever event started first must not decide the row, or the
  // pair swaps every time one side is replaced mid-overlap.
  const lyricLayout = { kind: 'source-order', sourceOrder: 0, verticalBand: 'top' } as const;
  const dialogueLayout = { kind: 'source-order', sourceOrder: 1, verticalBand: 'bottom' } as const;
  const dialogue = {
    startTime: 632.2,
    endTime: 634.8,
    text: '\u30e9\u30a4\u30d6\u3000\u3084\u3081\u3088\u3063\u304b',
    assLayout: dialogueLayout,
  };

  // Lyric started before the dialogue...
  assert.equal(
    resolvePrimarySubtitleText({
      liveText: '\u30e9\u30a4\u30d6\u3000\u3084\u3081\u3088\u3063\u304b\n\u6b4c\u8a5e\uff21',
      currentTimeSec: 632.5,
      cues: [
        { startTime: 629.5, endTime: 633.5, text: '\u6b4c\u8a5e\uff21', assLayout: lyricLayout },
        dialogue,
      ],
    }),
    '\u6b4c\u8a5e\uff21\n\n\u30e9\u30a4\u30d6\u3000\u3084\u3081\u3088\u3063\u304b',
  );
  // ...and the next lyric starts after it: the rows must not swap.
  assert.equal(
    resolvePrimarySubtitleText({
      liveText: '\u30e9\u30a4\u30d6\u3000\u3084\u3081\u3088\u3063\u304b\n\u6b4c\u8a5e\uff22',
      currentTimeSec: 633.8,
      cues: [
        dialogue,
        { startTime: 633.5, endTime: 637.0, text: '\u6b4c\u8a5e\uff22', assLayout: lyricLayout },
      ],
    }),
    '\u6b4c\u8a5e\uff22\n\n\u30e9\u30a4\u30d6\u3000\u3084\u3081\u3088\u3063\u304b',
  );
});

test('resolvePrimarySubtitleText puts an unreadable placement above bottom dialogue', () => {
  // Dialogue is the case that reliably declares a bottom alignment, so a cue whose
  // placement could not be read is more often a sign or song line. Keeping dialogue on
  // the bottom row means the line worth reading stays where the eye already is.
  assert.equal(
    resolvePrimarySubtitleText({
      liveText: '\u4e0b\u306e\u30bb\u30ea\u30d5\n\u4e0d\u660e\u306a\u884c',
      currentTimeSec: 2,
      cues: [
        {
          startTime: 1,
          endTime: 3,
          text: '\u4e0b\u306e\u30bb\u30ea\u30d5',
          assLayout: { kind: 'source-order', sourceOrder: 0, verticalBand: 'bottom' },
        },
        {
          startTime: 1.5,
          endTime: 3,
          text: '\u4e0d\u660e\u306a\u884c',
          assLayout: { kind: 'source-order', sourceOrder: 1 },
        },
      ],
    }),
    '\u4e0d\u660e\u306a\u884c\n\n\u4e0b\u306e\u30bb\u30ea\u30d5',
  );
});

test('resolvePrimarySubtitleText keeps source order when no cue declares a placement', () => {
  // SRT and websocket cues carry no layout at all: every cue ties, so the stable sort
  // must leave them exactly as the cue list had them.
  assert.equal(
    resolvePrimarySubtitleText({
      liveText: 'First line\nSecond line',
      currentTimeSec: 2,
      cues: [
        { startTime: 1, endTime: 3, text: 'First line' },
        { startTime: 1.5, endTime: 3, text: 'Second line' },
      ],
    }),
    'First line\n\nSecond line',
  );
});

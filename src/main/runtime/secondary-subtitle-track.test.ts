import assert from 'node:assert/strict';
import test from 'node:test';
import { parseSubtitleCues } from '../../core/services/subtitle-cue-parser';
import {
  createSecondarySubtitleTrackController,
  findActiveSubtitleText,
} from './secondary-subtitle-track';

test('findActiveSubtitleText combines unique simultaneous parsed cues', () => {
  assert.equal(
    findActiveSubtitleText(
      [
        { startTime: 1, endTime: 3, text: 'Your' },
        { startTime: 1, endTime: 3, text: 'Your' },
        { startTime: 1, endTime: 3, text: 'mosaic' },
      ],
      2,
    ),
    'Your\nmosaic',
  );
});

test('findActiveSubtitleText removes duplicate lines across multiline cues', () => {
  assert.equal(
    findActiveSubtitleText(
      [
        { startTime: 1, endTime: 3, text: 'First line\nSecond line' },
        { startTime: 1, endTime: 3, text: 'First line' },
      ],
      2,
    ),
    'First line\nSecond line',
  );
});

test('findActiveSubtitleText removes equivalent full-width duplicate lines', () => {
  assert.equal(
    findActiveSubtitleText(
      [
        { startTime: 1, endTime: 3, text: '真白～' },
        { startTime: 1, endTime: 3, text: '真白~' },
      ],
      2,
    ),
    '真白～',
  );
});

test('findActiveSubtitleText collapses whitespace variants of one ASS lyric', () => {
  assert.equal(
    findActiveSubtitleText(
      [
        { startTime: 1, endTime: 3, text: '少しだけ好きになる' },
        { startTime: 1, endTime: 3, text: '少しだけ 好きになる' },
        { startTime: 1, endTime: 3, text: '少しだけ　好きになる' },
      ],
      2,
    ),
    '少しだけ好きになる',
  );
});

test('parsed secondary text collapses a positioned sign that repeats dialogue without punctuation', () => {
  const ass = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 10,0:03:58.49,0:04:00.34,GJM_Main_1080p,Nar,0,0,0,,{\\i1}A question veiled as an insult!',
    'Dialogue: 1,0:03:58.59,0:04:00.34,iFanzSigns,,0,0,0,,{\\pos(960,75)}A question veiled as an insult',
  ].join('\n');
  const cues = parseSubtitleCues(ass, 'kaguya-s02e10.ass');

  assert.equal(findActiveSubtitleText(cues, 238.48), '');
  assert.equal(findActiveSubtitleText(cues, 238.5), 'A question veiled as an insult!');
  assert.equal(findActiveSubtitleText(cues, 239), 'A question veiled as an insult!');
  assert.equal(findActiveSubtitleText(cues, 240.34), '');
});

test('parsed secondary text drops a reconstructed grid of positioned sign fragments', () => {
  const signFragment = (text: string, x: number, y: number) =>
    `Dialogue: 1,0:00:01.00,0:00:03.00,Signs,,0,0,0,,{\\pos(${x},${y})\\t(0,100,\\fscx101)}${text}`;
  const ass = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 10,0:00:01.00,0:00:03.00,Default,Speaker,0,0,0,,Come on, wake up!',
    signFragment('Timetable', 1700, 150),
    signFragment('Mon', 1750, 230),
    signFragment('Tue', 1850, 230),
    signFragment('1', 1650, 320),
    signFragment('2', 1650, 390),
    signFragment('Civics', 1750, 320),
    signFragment('Math', 1850, 390),
    signFragment('PE', 1850, 460),
  ].join('\n');

  assert.equal(
    findActiveSubtitleText(parseSubtitleCues(ass, 'kaguya-s02e11.ass'), 2),
    'Come on, wake up!',
  );
});

test('parsed secondary lyrics keep explicit ASS vertical order when durations alternate', () => {
  const lyric = (options: { start: string; end: string; style: string; y: number; text: string }) =>
    `Dialogue: 0,0:00:${options.start},0:00:${options.end},${options.style},,0,0,0,fx,{\\move(100,${options.y},120,${options.y})\\t(0,200,\\fscx110)}${options.text}\\N{\\p1}m 0 0 l 0 5`;
  const ass = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    lyric({
      start: '01.00',
      end: '02.20',
      style: 'ed_romaji',
      y: 66,
      text: 'ima wo kakusarechau mae ni',
    }),
    lyric({
      start: '01.00',
      end: '02.00',
      style: 'ed_english',
      y: 1020,
      text: 'Before the present moment gets hidden away.',
    }),
    lyric({
      start: '03.00',
      end: '04.00',
      style: 'ed_romaji',
      y: 66,
      text: 'ame mitai ni hikatteru',
    }),
    lyric({
      start: '03.00',
      end: '04.20',
      style: 'ed_english',
      y: 1020,
      text: 'Is shining like rain.',
    }),
  ].join('\n');
  const cues = parseSubtitleCues(ass, 'polar-opposites-s01e08.ass');

  assert.equal(
    findActiveSubtitleText(cues, 1.5),
    'ima wo kakusarechau mae ni\nBefore the present moment gets hidden away.',
  );
  assert.equal(findActiveSubtitleText(cues, 3.5), 'ame mitai ni hikatteru\nIs shining like rain.');
});

test('unpositioned secondary lyrics fall back to ASS source order', () => {
  const ass = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:02.20,ED Romaji,,0,0,0,,ima wo kakusarechau mae ni',
    'Dialogue: 0,0:00:01.00,0:00:02.00,ED English,,0,0,0,,Before the present moment gets hidden away.',
  ].join('\n');

  assert.equal(
    findActiveSubtitleText(parseSubtitleCues(ass, 'ending.ass'), 1.5),
    'ima wo kakusarechau mae ni\nBefore the present moment gets hidden away.',
  );
});

test('findActiveSubtitleText keeps a canonical ASS cue for its generated animation span', () => {
  const poof = {
    startTime: 1110.67,
    endTime: 1110.71,
    text: 'POOF',
    source: 'canonical-ass' as const,
    animationStartTime: 1110.67,
    animationEndTime: 1111.59,
  };

  assert.equal(findActiveSubtitleText([poof], 1111.58), 'POOF');
  assert.equal(findActiveSubtitleText([poof], 1111.59), '');
});

test('findActiveSubtitleText advances when the next canonical lyric animation starts', () => {
  const cues = [
    {
      startTime: 121.73,
      endTime: 124.1,
      text: 'Torn at the seams, a sound pours out',
      source: 'canonical-ass' as const,
      animationStartTime: 121.4,
      animationEndTime: 124.1,
    },
    {
      startTime: 124.13,
      endTime: 126.38,
      text: 'It’s silent, yet spreads all around',
      source: 'canonical-ass' as const,
      animationStartTime: 123.8,
      animationEndTime: 126.38,
    },
  ];

  assert.equal(findActiveSubtitleText(cues, 123.79), cues[0]!.text);
  assert.equal(findActiveSubtitleText(cues, 123.8), cues[1]!.text);
});

test('ASS fragment karaoke stays separated by style with authored word spacing', () => {
  const lineEvents = (
    style: string,
    fragments: readonly string[],
    y: number,
    baseTime = 1,
  ): string[] => {
    const events: string[] = [];
    for (const layer of [0, 1]) {
      fragments.forEach((fragment, index) => {
        const x = 100 + index * 40;
        const start = (baseTime + index * 0.25).toFixed(2).padStart(5, '0');
        const end = (baseTime + 3 + index * 0.2).toFixed(2).padStart(5, '0');
        events.push(
          `Dialogue: ${layer},0:00:${start},0:00:${end},${style},,0,0,0,,{\\pos(${x},${y})\\t(0,200,\\fscx110)}${fragment}\\N{\\p1}m 0 0 l 0 10`,
        );
      });
    }
    return events;
  };
  const ass = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    ...lineEvents('ed_romaji', ['ji', 'gu', 'za', 'gu ', 'na', 'mi'], 70),
    ...lineEvents('ed_english', ['Pas', 'si', 'ng ', 'thro', 'u', 'gh '], 110),
    ...lineEvents('op_english', ['I', 'want', 'to', 'go'], 110, 7),
  ].join('\n');

  assert.equal(
    findActiveSubtitleText(parseSubtitleCues(ass, 'ending.ass'), 2.5),
    'jiguzagu nami\nPassing through',
  );
  // Some generated scripts discard spaces and retain only positioned chunks. Joining
  // without invented separators avoids turning one word into spaced syllables.
  assert.equal(findActiveSubtitleText(parseSubtitleCues(ass, 'ending.ass'), 8.5), 'Iwanttogo');
});

test('ASS fragment karaoke preserves word spaces authored at event boundaries', () => {
  const fragments = [
    'The ',
    'shoot',
    'ing ',
    'stars ',
    'arc',
    'ing ',
    'across ',
    'the ',
    'sky ',
    'I ',
    'wish ',
    'upon,',
  ];
  const events: string[] = [];
  for (const layer of [0, 1]) {
    fragments.forEach((fragment, index) => {
      events.push(
        `Dialogue: ${layer},0:00:01.00,0:00:04.00,op_english,,0,0,0,,{\\pos(${100 + index * 40},110)\\t(0,200,\\fscx110)}${fragment}`,
      );
    });
  }
  const ass = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    ...events,
  ].join('\n');

  assert.equal(
    findActiveSubtitleText(parseSubtitleCues(ass, 'bravern-s01e10.ass'), 2),
    'The shooting stars arcing across the sky I wish upon,',
  );
});

test('findActiveSubtitleText keeps a complete reconstructed line over entrance fragments', () => {
  const current = {
    startTime: 1,
    endTime: 4,
    text: 'Complete current line',
    source: 'reconstructed-ass' as const,
    assStyle: 'op_english',
  };
  const nextEntrance = {
    startTime: 3.8,
    endTime: 4.2,
    text: 'Ne',
    source: 'reconstructed-ass' as const,
    assStyle: 'op_english',
  };
  const nextLine = {
    startTime: 4,
    endTime: 7,
    text: 'Next complete line',
    source: 'reconstructed-ass' as const,
    assStyle: 'op_english',
  };

  assert.equal(findActiveSubtitleText([current, nextEntrance], 3.9), current.text);
  assert.equal(findActiveSubtitleText([current, nextEntrance, nextLine], 4.1), nextLine.text);
});

test('secondary track controller parses the selected ASS file before publishing', async () => {
  const broadcasts: string[] = [];
  let currentText = '';
  const resolverInputs: Array<{ allowSelectedFallback?: boolean }> = [];
  const ass = `[Events]
Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
Dialogue: 0,0:00:01.00,0:00:03.00,Sign,,0,0,0,,Your
Dialogue: 1,0:00:01.00,0:00:03.00,Sign,,0,0,0,,Your
Dialogue: 2,0:00:01.00,0:00:03.00,Sign,,0,0,0,,Your
Dialogue: 3,0:00:01.00,0:00:03.00,Sign,,0,0,0,,Your
Dialogue: 4,0:00:01.00,0:00:03.00,Sign,,0,0,0,,mosaic`;
  const controller = createSecondarySubtitleTrackController({
    getMpvClient: () => ({
      connected: true,
      requestProperty: async (name) => {
        if (name === 'secondary-sid') return 2;
        if (name === 'track-list') return [{ type: 'sub', id: 2 }];
        if (name === 'path') return '/media/video.mkv';
        if (name === 'secondary-sub-delay') return 0;
        return null;
      },
    }),
    getCurrentTimePos: () => 2,
    resolveSubtitleSource: async (input) => {
      resolverInputs.push(input);
      return { path: '/subs/english.ass', sourceKey: '/subs/english.ass' };
    },
    loadSubtitleSourceText: async () => ass,
    parseSubtitleCues,
    setCurrentSecondaryText: (text) => {
      currentText = text;
    },
    broadcastSecondaryText: (text) => broadcasts.push(text),
  });

  await controller.refresh();
  controller.handleLiveText('Your\nYour\nYour\nYour\nmosaic');

  assert.equal(resolverInputs[0]?.allowSelectedFallback, false);
  assert.equal(currentText, 'Your\nmosaic');
  assert.deepEqual(broadcasts, ['Your\nmosaic']);
});

test('secondary track controller follows parsed cue timing and subtitle delay', async () => {
  const broadcasts: string[] = [];
  let time = 2.25;
  const controller = createSecondarySubtitleTrackController({
    getMpvClient: () => ({
      connected: true,
      requestProperty: async (name) => {
        if (name === 'secondary-sid') return 2;
        if (name === 'track-list') return [{ type: 'sub', id: 2 }];
        if (name === 'path') return '/media/video.mkv';
        if (name === 'secondary-sub-delay') return 0.5;
        return null;
      },
    }),
    getCurrentTimePos: () => time,
    resolveSubtitleSource: async () => ({ path: '/subs/english.srt', sourceKey: 'english' }),
    loadSubtitleSourceText: async () => '',
    parseSubtitleCues: () => [
      { startTime: 1, endTime: 2, text: 'first' },
      { startTime: 2, endTime: 3, text: 'second' },
    ],
    setCurrentSecondaryText: () => {},
    broadcastSecondaryText: (text) => broadcasts.push(text),
  });

  await controller.refresh();
  controller.handleDelayChange(0);
  time = 3.25;
  controller.handleTimePos(time);

  assert.deepEqual(broadcasts, ['first', 'second', '']);
});

test('secondary track controller clears old parsed text immediately on a track change', async () => {
  const broadcasts: string[] = [];
  let currentText = '';
  const controller = createSecondarySubtitleTrackController({
    getMpvClient: () => ({
      connected: true,
      requestProperty: async (name) => {
        if (name === 'secondary-sid') return 2;
        if (name === 'track-list') return [{ type: 'sub', id: 2, external: true }];
        if (name === 'path') return '/media/video.mkv';
        if (name === 'secondary-sub-delay') return 0;
        return null;
      },
    }),
    getCurrentTimePos: () => 2,
    resolveSubtitleSource: async () => ({ path: '/subs/old.ass', sourceKey: 'old' }),
    loadSubtitleSourceText: async () => '',
    parseSubtitleCues: () => [{ startTime: 1, endTime: 3, text: 'old parsed text' }],
    setCurrentSecondaryText: (text) => {
      currentText = text;
    },
    broadcastSecondaryText: (text) => broadcasts.push(text),
  });

  await controller.refresh();
  controller.handleTrackChange();
  controller.handleLiveText('new live text');

  assert.equal(currentText, 'new live text');
  assert.deepEqual(broadcasts, ['old parsed text', '', 'new live text']);
});

test('secondary track controller falls back to live mpv text without a readable source', async () => {
  const broadcasts: string[] = [];
  let currentText = '';
  const controller = createSecondarySubtitleTrackController({
    getMpvClient: () => ({
      connected: true,
      requestProperty: async (name) => {
        if (name === 'secondary-sid') return 'no';
        if (name === 'path') return '/media/video.mkv';
        return null;
      },
    }),
    getCurrentTimePos: () => 2,
    resolveSubtitleSource: async () => null,
    loadSubtitleSourceText: async () => '',
    parseSubtitleCues: () => [],
    setCurrentSecondaryText: (text) => {
      currentText = text;
    },
    broadcastSecondaryText: (text) => broadcasts.push(text),
  });

  controller.handleLiveText('live fallback');
  await controller.refresh();

  assert.equal(currentText, 'live fallback');
  assert.deepEqual(broadcasts, ['live fallback']);
});

test('secondary ASS live fallback drops malformed control debris', async () => {
  const broadcasts: string[] = [];
  const controller = createSecondarySubtitleTrackController({
    getMpvClient: () => ({
      connected: true,
      requestProperty: async (name) => {
        if (name === 'secondary-sid') return 2;
        if (name === 'track-list') return [{ type: 'sub', id: 2 }];
        if (name === 'path') return '/media/video.mkv';
        return null;
      },
    }),
    getCurrentTimePos: () => 2,
    resolveSubtitleSource: async () => ({ path: '/subs/english.ass', sourceKey: 'english' }),
    loadSubtitleSourceText: async () => '',
    parseSubtitleCues: () => [],
    setCurrentSecondaryText: () => {},
    broadcastSecondaryText: (text) => broadcasts.push(text),
  });

  await controller.refresh();
  broadcasts.length = 0;
  controller.handleLiveText('Visible line\n\\\n{\\fr0');

  assert.deepEqual(broadcasts, ['Visible line']);
});

test('secondary SRT live fallback preserves text that resembles ASS control debris', async () => {
  const broadcasts: string[] = [];
  const controller = createSecondarySubtitleTrackController({
    getMpvClient: () => ({
      connected: true,
      requestProperty: async (name) => {
        if (name === 'secondary-sid') return 2;
        if (name === 'track-list') return [{ type: 'sub', id: 2 }];
        if (name === 'path') return '/media/video.mkv';
        return null;
      },
    }),
    getCurrentTimePos: () => 2,
    resolveSubtitleSource: async () => ({ path: '/subs/english.srt', sourceKey: 'english' }),
    loadSubtitleSourceText: async () => '',
    parseSubtitleCues: () => [],
    setCurrentSecondaryText: () => {},
    broadcastSecondaryText: (text) => broadcasts.push(text),
  });

  await controller.refresh();
  broadcasts.length = 0;
  controller.handleLiveText('Visible line\n\\\n{\\fr0');

  assert.deepEqual(broadcasts, ['Visible line\n\\\n{\\fr0']);
});

test('secondary disconnect clears stale ASS fallback sanitization state', async () => {
  let connected = true;
  const broadcasts: string[] = [];
  const controller = createSecondarySubtitleTrackController({
    getMpvClient: () => ({
      connected,
      requestProperty: async (name) => {
        if (name === 'secondary-sid') return 2;
        if (name === 'track-list') return [{ type: 'sub', id: 2 }];
        if (name === 'path') return '/media/video.mkv';
        return null;
      },
    }),
    getCurrentTimePos: () => 2,
    resolveSubtitleSource: async () => ({ path: '/subs/english.ass', sourceKey: 'english' }),
    loadSubtitleSourceText: async () => '',
    parseSubtitleCues: () => [],
    setCurrentSecondaryText: () => {},
    broadcastSecondaryText: (text) => broadcasts.push(text),
  });

  await controller.refresh();
  connected = false;
  await controller.refresh();
  broadcasts.length = 0;
  controller.handleLiveText('Visible line\n\\\n{\\fr0');

  assert.deepEqual(broadcasts, ['Visible line\n\\\n{\\fr0']);
});

test('secondary source refresh failure clears stale ASS fallback sanitization state', async () => {
  let resolveCalls = 0;
  const broadcasts: string[] = [];
  const controller = createSecondarySubtitleTrackController({
    getMpvClient: () => ({
      connected: true,
      requestProperty: async (name) => {
        if (name === 'secondary-sid') return 2;
        if (name === 'track-list') return [{ type: 'sub', id: 2 }];
        if (name === 'path') return '/media/video.mkv';
        return null;
      },
    }),
    getCurrentTimePos: () => 2,
    resolveSubtitleSource: async () => {
      resolveCalls += 1;
      if (resolveCalls === 1) {
        return { path: '/subs/english.ass', sourceKey: 'english' };
      }
      throw new Error('source refresh failed');
    },
    loadSubtitleSourceText: async () => '',
    parseSubtitleCues: () => [],
    setCurrentSecondaryText: () => {},
    broadcastSecondaryText: (text) => broadcasts.push(text),
  });

  await controller.refresh();
  await controller.refresh();
  broadcasts.length = 0;
  controller.handleLiveText('Visible line\n\\\n{\\fr0');

  assert.deepEqual(broadcasts, ['Visible line\n\\\n{\\fr0']);
});

test('secondary track controller reuses parsed cues for an unchanged embedded track', async () => {
  let resolveCalls = 0;
  let parseCalls = 0;
  const controller = createSecondarySubtitleTrackController({
    getMpvClient: () => ({
      connected: true,
      requestProperty: async (name) => {
        if (name === 'secondary-sid') return 2;
        if (name === 'track-list') {
          return [{ type: 'sub', id: 2, external: false, 'ff-index': 3 }];
        }
        if (name === 'path') return '/media/video.mkv';
        if (name === 'secondary-sub-delay') return 0;
        return null;
      },
    }),
    getCurrentTimePos: () => 2,
    resolveSubtitleSource: async () => {
      resolveCalls += 1;
      return { path: `/tmp/extracted-${resolveCalls}.ass`, sourceKey: 'embedded-track-2' };
    },
    loadSubtitleSourceText: async () => '',
    parseSubtitleCues: () => {
      parseCalls += 1;
      return [{ startTime: 1, endTime: 3, text: 'parsed' }];
    },
    setCurrentSecondaryText: () => {},
    broadcastSecondaryText: () => {},
  });

  await controller.refresh();
  await controller.refresh();

  assert.equal(resolveCalls, 1);
  assert.equal(parseCalls, 1);
});

test('secondary track controller ignores and cleans up a refresh invalidated by reset', async () => {
  const broadcasts: string[] = [];
  let notifyResolveStarted: (() => void) | undefined;
  let releaseResolve: (() => void) | undefined;
  let cleanupCalls = 0;
  let parseCalls = 0;
  const resolveStarted = new Promise<void>((resolve) => {
    notifyResolveStarted = resolve;
  });
  const resolveGate = new Promise<void>((resolve) => {
    releaseResolve = resolve;
  });
  const controller = createSecondarySubtitleTrackController({
    getMpvClient: () => ({
      connected: true,
      requestProperty: async (name) => {
        if (name === 'secondary-sid') return 2;
        if (name === 'track-list') return [{ type: 'sub', id: 2, external: true }];
        if (name === 'path') return '/media/video.mkv';
        if (name === 'secondary-sub-delay') return 0;
        return null;
      },
    }),
    getCurrentTimePos: () => 2,
    resolveSubtitleSource: async () => {
      notifyResolveStarted?.();
      await resolveGate;
      return {
        path: '/subs/secondary.ass',
        sourceKey: 'secondary',
        cleanup: async () => {
          cleanupCalls += 1;
        },
      };
    },
    loadSubtitleSourceText: async () => '',
    parseSubtitleCues: () => {
      parseCalls += 1;
      return [{ startTime: 1, endTime: 3, text: 'stale' }];
    },
    setCurrentSecondaryText: () => {},
    broadcastSecondaryText: (text) => broadcasts.push(text),
  });

  const refresh = controller.refresh();
  await resolveStarted;
  controller.reset();
  releaseResolve?.();
  await refresh;

  assert.deepEqual(broadcasts, ['']);
  assert.equal(parseCalls, 0);
  assert.equal(cleanupCalls, 1);
});

test('secondary live fallback suppresses a per-glyph typesetting wall', async () => {
  let currentText = '';
  const controller = createSecondarySubtitleTrackController({
    getMpvClient: () => ({
      connected: true,
      requestProperty: async (name) => {
        if (name === 'secondary-sid') return 2;
        if (name === 'track-list') return [{ type: 'sub', id: 2 }];
        if (name === 'path') return '/mnt/nas/video.mkv';
        if (name === 'secondary-sub-delay') return 0;
        return null;
      },
    }),
    getCurrentTimePos: () => 1355,
    // Network-mounted media: embedded extraction is skipped, so no parsed cues exist.
    resolveSubtitleSource: async () => null,
    loadSubtitleSourceText: async () => '',
    parseSubtitleCues,
    setCurrentSecondaryText: (text) => {
      currentText = text;
    },
    broadcastSecondaryText: () => {},
  });

  await controller.refresh();
  const wall = [...'wansdumretoikhI'].join('\n');
  controller.handleLiveText(`${wall}\ntai`);
  assert.equal(currentText, '');

  controller.handleLiveText(`${wall}\nそれよりも　ノート…`);
  assert.equal(currentText, 'それよりも　ノート…');
});

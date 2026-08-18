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

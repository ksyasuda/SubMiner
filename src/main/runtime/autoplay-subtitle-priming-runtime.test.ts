import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createAutoplaySubtitlePrimingRuntime,
  setMpvCurrentSecondarySubText,
} from './autoplay-subtitle-priming-runtime';

test('setMpvCurrentSecondarySubText uses client setter when available', () => {
  const calls: string[] = [];
  const client = {
    currentSecondarySubText: '',
    setCurrentSecondarySubText: (text: string) => {
      calls.push(text);
    },
  };

  setMpvCurrentSecondarySubText(client, 'secondary');

  assert.deepEqual(calls, ['secondary']);
  assert.equal(client.currentSecondarySubText, '');
});

test('setMpvCurrentSecondarySubText updates client property when setter is unavailable', () => {
  const client = {
    currentSecondarySubText: '',
  };

  setMpvCurrentSecondarySubText(client, 'secondary');

  assert.equal(client.currentSecondarySubText, 'secondary');
});

test('scheduleSubtitlePrefetchRefresh logs refresh failures from timer callback', async () => {
  const logs: string[] = [];
  const runtime = createAutoplaySubtitlePrimingRuntime({
    getCurrentMediaPath: () => null,
    getMpvClient: () => null,
    setCurrentSubText: () => {},
    getCurrentSubText: () => '',
    getCurrentSubtitleData: () => null,
    getActiveParsedSubtitleCues: () => [],
    setActiveParsedSubtitleMediaPath: () => {},
    subtitleProcessingController: {
      consumeCachedSubtitle: () => null,
      onSubtitleChange: () => {},
      refreshCurrentSubtitle: () => {},
    },
    emitSubtitlePayload: () => {},
    getSubtitlePrefetchService: () => null,
    getLastObservedTimePos: () => 0,
    getVisibleOverlayVisible: () => false,
    emitSecondarySubtitle: () => {},
    initSubtitlePrefetch: async () => {},
    refreshSubtitlePrefetchFromActiveTrack: async () => {
      throw new Error('refresh failed');
    },
    logDebug: (message) => logs.push(message),
  });

  runtime.scheduleSubtitlePrefetchRefresh(0);
  await new Promise((resolve) => setTimeout(resolve, 5));

  assert.deepEqual(logs, [
    '[autoplay-subtitle-prime] subtitle prefetch refresh failed: refresh failed',
  ]);
});

test('primeCurrentSubtitleForAutoplay refreshes active subtitle cues when mpv sub-text is empty', async () => {
  const calls: string[] = [];
  let currentSubText = '';
  let activeParsedSubtitleCues: Array<{ startTime: number; endTime: number; text: string }> = [];
  const mediaPath = '/media/video.mkv';

  const runtime = createAutoplaySubtitlePrimingRuntime({
    getCurrentMediaPath: () => mediaPath,
    getMpvClient: () => ({
      connected: true,
      currentVideoPath: mediaPath,
      requestProperty: async (name) => {
        calls.push(`request:${name}`);
        if (name === 'sub-text') return '';
        if (name === 'time-pos') return 12;
        return null;
      },
    }),
    setCurrentSubText: (text) => {
      currentSubText = text;
      calls.push(`set:${text}`);
    },
    getCurrentSubText: () => currentSubText,
    getCurrentSubtitleData: () => null,
    getActiveParsedSubtitleCues: () => activeParsedSubtitleCues,
    setActiveParsedSubtitleMediaPath: () => {},
    subtitleProcessingController: {
      consumeCachedSubtitle: () => null,
      onSubtitleChange: (text) => calls.push(`change:${text}`),
      refreshCurrentSubtitle: (text) => calls.push(`refresh:${text ?? ''}`),
    },
    emitSubtitlePayload: (payload) => calls.push(`emit:${payload.text}`),
    getSubtitlePrefetchService: () => ({
      pause: () => calls.push('prefetch:pause'),
      onSeek: (timePos) => calls.push(`prefetch:seek:${timePos}`),
    }),
    getLastObservedTimePos: () => 12,
    getVisibleOverlayVisible: () => true,
    emitSecondarySubtitle: () => {},
    initSubtitlePrefetch: async () => {},
    refreshSubtitlePrefetchFromActiveTrack: async () => {
      calls.push('refresh-active-track');
      activeParsedSubtitleCues = [{ startTime: 10, endTime: 20, text: '起動字幕' }];
    },
    logDebug: (message) => calls.push(`debug:${message}`),
  });

  await runtime.primeCurrentSubtitleForAutoplay(mediaPath);

  assert.deepEqual(calls, [
    'request:sub-text',
    'refresh-active-track',
    'request:time-pos',
    'set:起動字幕',
    'prefetch:pause',
    'emit:起動字幕',
    'change:起動字幕',
  ]);
});

test('primeCurrentSubtitleForAutoplay emits raw first paint on cache miss before tokenization', async () => {
  const calls: string[] = [];
  let currentSubText = '';
  const mediaPath = '/media/video.mkv';

  const runtime = createAutoplaySubtitlePrimingRuntime({
    getCurrentMediaPath: () => mediaPath,
    getMpvClient: () => ({
      connected: true,
      currentVideoPath: mediaPath,
      requestProperty: async (name) => {
        calls.push(`request:${name}`);
        if (name === 'sub-text') return '起動字幕';
        return null;
      },
    }),
    setCurrentSubText: (text) => {
      currentSubText = text;
      calls.push(`set:${text}`);
    },
    getCurrentSubText: () => currentSubText,
    getCurrentSubtitleData: () => null,
    getActiveParsedSubtitleCues: () => [],
    setActiveParsedSubtitleMediaPath: () => {},
    subtitleProcessingController: {
      consumeCachedSubtitle: () => null,
      onSubtitleChange: (text) => calls.push(`change:${text}`),
      refreshCurrentSubtitle: (text) => calls.push(`refresh:${text ?? ''}`),
    },
    emitSubtitlePayload: (payload) => calls.push(`emit:${payload.text}`),
    getSubtitlePrefetchService: () => ({
      pause: () => calls.push('prefetch:pause'),
      onSeek: (timePos) => calls.push(`prefetch:seek:${timePos}`),
    }),
    getLastObservedTimePos: () => 12,
    getVisibleOverlayVisible: () => true,
    emitSecondarySubtitle: () => {},
    initSubtitlePrefetch: async () => {},
    refreshSubtitlePrefetchFromActiveTrack: async () => {
      calls.push('refresh-active-track');
    },
    logDebug: (message) => calls.push(`debug:${message}`),
  });

  await runtime.primeCurrentSubtitleForAutoplay(mediaPath);

  assert.deepEqual(calls, [
    'request:sub-text',
    'set:起動字幕',
    'prefetch:pause',
    'emit:起動字幕',
    'change:起動字幕',
  ]);
});

import assert from 'node:assert/strict';
import test from 'node:test';
import { parseSubtitleCues } from '../../core/services/subtitle-cue-parser';
import { createSubtitleProcessingController } from '../../core/services/subtitle-processing-controller';
import type { SubtitleData } from '../../types';
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
      onSubtitleChange: () => true,
      refreshCurrentSubtitle: () => true,
      notePlainSubtitleEmitted: () => {},
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
      onSubtitleChange: (text) => {
        calls.push(`change:${text}`);
        return true;
      },
      refreshCurrentSubtitle: (text) => {
        calls.push(`refresh:${text ?? ''}`);
        return true;
      },
      notePlainSubtitleEmitted: () => {},
    },
    emitSubtitlePayload: (payload, options) =>
      calls.push(`emit:${payload.text}:resume=${options?.resumePrefetch !== false}`),
    getSubtitlePrefetchService: () => ({
      pause: () => {
        calls.push('prefetch:pause');
      },
      resume: () => {
        calls.push('prefetch:resume');
      },
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
    'emit:起動字幕:resume=false',
    // Uncached priming refreshes rather than announcing a change, so an
    // invalidated-but-unchanged line is still re-tokenized.
    'refresh:起動字幕',
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
      onSubtitleChange: (text) => {
        calls.push(`change:${text}`);
        return true;
      },
      refreshCurrentSubtitle: (text) => {
        calls.push(`refresh:${text ?? ''}`);
        return true;
      },
      notePlainSubtitleEmitted: () => {},
    },
    emitSubtitlePayload: (payload, options) =>
      calls.push(`emit:${payload.text}:resume=${options?.resumePrefetch !== false}`),
    getSubtitlePrefetchService: () => ({
      pause: () => {
        calls.push('prefetch:pause');
      },
      resume: () => {
        calls.push('prefetch:resume');
      },
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
    'emit:起動字幕:resume=false',
    // Uncached priming refreshes rather than announcing a change, so an
    // invalidated-but-unchanged line is still re-tokenized.
    'refresh:起動字幕',
  ]);
});

test('parsed cues replace a duplicate raw autoplay subtitle that was already primed', async () => {
  const rawText = 'ジグザグな道を抜け\nジグザグな道を抜け';
  const correctedText = 'ジグザグな道を抜け';
  const mediaPath = '/media/video.mkv';
  let currentSubText = '';
  const emitted: string[] = [];
  const client = {
    connected: true,
    currentVideoPath: mediaPath,
    currentTimePos: 90,
    currentSubText: rawText,
    requestProperty: async (name: string) => {
      if (name === 'sub-text') return rawText;
      if (name === 'time-pos') return 90;
      return null;
    },
  };
  const cues = parseSubtitleCues(
    [
      '[Events]',
      'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
      `Dialogue: 1,0:01:29.00,0:01:32.00,EDJP,,0,0,0,,${correctedText}`,
      `Dialogue: 0,0:01:29.00,0:01:32.00,EDJP,,0,0,0,,${correctedText}`,
    ].join('\n'),
    'startup-ending.ass',
  );
  let activeCues = cues.slice(0, 0);
  const runtime = createAutoplaySubtitlePrimingRuntime({
    getCurrentMediaPath: () => mediaPath,
    getMpvClient: () => client,
    setCurrentSubText: (text) => {
      currentSubText = text;
    },
    getCurrentSubText: () => currentSubText,
    getCurrentSubtitleData: () => null,
    getActiveParsedSubtitleCues: () => activeCues,
    setActiveParsedSubtitleMediaPath: () => {},
    subtitleProcessingController: {
      consumeCachedSubtitle: () => null,
      onSubtitleChange: () => true,
      refreshCurrentSubtitle: () => true,
      notePlainSubtitleEmitted: () => {},
    },
    emitSubtitlePayload: (payload) => emitted.push(payload.text),
    getSubtitlePrefetchService: () => null,
    getLastObservedTimePos: () => 90,
    getVisibleOverlayVisible: () => true,
    emitSecondarySubtitle: () => {},
    initSubtitlePrefetch: async () => {},
    refreshSubtitlePrefetchFromActiveTrack: async () => {},
    logDebug: () => {},
  });

  await runtime.primeCurrentSubtitleForAutoplay(mediaPath);
  assert.equal(currentSubText, rawText);

  activeCues = cues;
  await runtime.primeAutoplaySubtitleFromParsedCues(mediaPath, cues);

  assert.equal(currentSubText, correctedText);
  assert.deepEqual(emitted, [rawText, correctedText]);
});

// Driven by the real processing controller rather than a stub: the failure this
// covers is a disagreement between the priming path and the controller's own
// staleness rules, which a hand-written stub cannot reproduce.
function createPrimingRuntimeWithRealController(options: {
  text: string;
  calls: string[];
  onTokenize: () => void;
  tokenize?: (text: string) => SubtitleData | null | Promise<SubtitleData | null>;
  cacheLimit?: number;
}) {
  const { text, calls } = options;
  let currentSubText = '';
  let currentSubtitleData: SubtitleData | null = null;
  const mediaPath = '/media/video.mkv';

  const prefetchService = {
    pause: () => calls.push('prefetch:pause'),
    resume: () => calls.push('prefetch:resume'),
  };
  // Mirrors main.ts emitSubtitlePayload: an emit resumes prefetching unless it
  // is explicitly marked as not the end of the work for the line, and every
  // controller emit is so marked.
  const emitSubtitlePayload = (
    payload: SubtitleData,
    emitOptions?: { resumePrefetch?: boolean },
  ): void => {
    currentSubtitleData = payload;
    calls.push(
      emitOptions?.resumePrefetch === false
        ? `emit-raw:${payload.text}`
        : `emit-direct:${payload.text}`,
    );
    if (emitOptions?.resumePrefetch !== false) {
      prefetchService.resume();
    }
  };

  const subtitleProcessingController = createSubtitleProcessingController({
    tokenizeSubtitle: async (subtitleText) => {
      options.onTokenize();
      return options.tokenize ? options.tokenize(subtitleText) : { text: subtitleText, tokens: [] };
    },
    // main.ts routes controller emits through emitSubtitlePayload with
    // resumePrefetch: false, so they never release the pause on their own.
    emitSubtitle: (payload) => {
      currentSubtitleData = payload;
      calls.push(`emit:${payload.text}:tokens=${payload.tokens === null ? 'none' : 'yes'}`);
    },
    onProcessingSettled: () => {
      prefetchService.resume();
    },
    ...(options.cacheLimit === undefined ? {} : { cacheLimit: options.cacheLimit }),
  });

  const runtime = createAutoplaySubtitlePrimingRuntime({
    getCurrentMediaPath: () => mediaPath,
    getMpvClient: () => ({
      connected: true,
      currentVideoPath: mediaPath,
      requestProperty: async (name) => (name === 'sub-text' ? text : null),
    }),
    setCurrentSubText: (value) => {
      currentSubText = value;
    },
    getCurrentSubText: () => currentSubText,
    getCurrentSubtitleData: () => currentSubtitleData,
    getActiveParsedSubtitleCues: () => [],
    setActiveParsedSubtitleMediaPath: () => {},
    subtitleProcessingController,
    emitSubtitlePayload,
    getSubtitlePrefetchService: () => prefetchService,
    getLastObservedTimePos: () => 12,
    getVisibleOverlayVisible: () => true,
    emitSecondarySubtitle: () => {},
    initSubtitlePrefetch: async () => {},
    refreshSubtitlePrefetchFromActiveTrack: async () => {},
    logDebug: () => {},
  });

  return { runtime, subtitleProcessingController, mediaPath };
}

test('primeCurrentSubtitleForAutoplay re-tokenizes text whose cached annotation was invalidated', async () => {
  const calls: string[] = [];
  let tokenizations = 0;
  const text = '起動字幕';
  const { runtime, subtitleProcessingController, mediaPath } =
    createPrimingRuntimeWithRealController({
      text,
      calls,
      onTokenize: () => {
        tokenizations += 1;
      },
    });

  // The line was already tokenized and cached during normal playback.
  subtitleProcessingController.onSubtitleChange(text);
  await new Promise((resolve) => setTimeout(resolve, 0));
  const tokenizationsBeforeInvalidation = tokenizations;

  // Mining a card drops every cached tokenization.
  subtitleProcessingController.invalidateTokenizationCache();
  calls.length = 0;

  await runtime.primeCurrentSubtitleForAutoplay(mediaPath);
  await new Promise((resolve) => setTimeout(resolve, 0));

  // The cache miss must schedule fresh work, or the line stays unannotated for
  // as long as it is on screen.
  assert.equal(
    tokenizations,
    tokenizationsBeforeInvalidation + 1,
    'expected the invalidated subtitle to be tokenized again',
  );
  assert.ok(
    calls.includes(`emit:${text}:tokens=yes`),
    `expected an annotated emit, saw ${JSON.stringify(calls)}`,
  );
});

test('primeCurrentSubtitleForAutoplay releases the prefetch pause when nothing is scheduled', async () => {
  const calls: string[] = [];
  const text = '起動字幕';
  const { runtime, subtitleProcessingController, mediaPath } =
    createPrimingRuntimeWithRealController({
      text,
      calls,
      onTokenize: () => {},
      cacheLimit: 1,
    });

  // Emitted at the current cache generation, then evicted from the one-entry
  // cache: priming misses the cache but the controller has nothing to redo, so
  // no emit is coming and the pause must be released here.
  subtitleProcessingController.onSubtitleChange(text);
  await new Promise((resolve) => setTimeout(resolve, 0));
  subtitleProcessingController.preCacheTokenization('別の字幕', {
    text: '別の字幕',
    tokens: [],
  });
  calls.length = 0;

  await runtime.primeCurrentSubtitleForAutoplay(mediaPath);
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.deepEqual(calls, ['prefetch:pause', `emit-raw:${text}`, 'prefetch:resume']);
});

test('primeCurrentSubtitleForAutoplay releases the prefetch pause when tokenization emits nothing', async () => {
  const calls: string[] = [];
  const text = '起動字幕';
  const { runtime, subtitleProcessingController, mediaPath } =
    createPrimingRuntimeWithRealController({
      text,
      calls,
      onTokenize: () => {},
      // Transient tokenizer failure: the controller falls back to plain text it
      // has already shown, so it suppresses the emit entirely.
      tokenize: () => null,
    });

  subtitleProcessingController.onSubtitleChange(text);
  await new Promise((resolve) => setTimeout(resolve, 0));
  subtitleProcessingController.invalidateTokenizationCache();
  calls.length = 0;

  await runtime.primeCurrentSubtitleForAutoplay(mediaPath);
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.ok(
    !calls.some((call) => call.startsWith('emit:')),
    `expected no controller emit, saw ${JSON.stringify(calls)}`,
  );
  assert.equal(
    calls.filter((call) => call === 'prefetch:resume').length,
    1,
    `expected the prefetch pause to be released, saw ${JSON.stringify(calls)}`,
  );
});

test('prefetch stays paused until tokenization of an uncached line completes', async () => {
  const calls: string[] = [];
  const text = '起動字幕';
  let finishTokenization = (): void => {};
  const tokenizationGate = new Promise<void>((resolve) => {
    finishTokenization = resolve;
  });
  const { runtime, mediaPath } = createPrimingRuntimeWithRealController({
    text,
    calls,
    onTokenize: () => {},
    tokenize: async (subtitleText) => {
      await tokenizationGate;
      return { text: subtitleText, tokens: [] };
    },
  });

  // Driven through the priming path, which is what takes the pause out.
  await runtime.primeCurrentSubtitleForAutoplay(mediaPath);
  await new Promise((resolve) => setTimeout(resolve, 0));

  // Neither the priming emit nor the controller's provisional plain emit may
  // release the pause: the expensive scan is still ahead of them and would
  // compete with prefetching for the parser.
  // One plain payload, not two: priming paints it and tells the controller, so
  // the controller goes straight for the tokenized one. And it does not resume
  // prefetch, because the expensive scan is still ahead of it.
  assert.deepEqual(calls, ['prefetch:pause', `emit-raw:${text}`]);

  finishTokenization();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.deepEqual(calls, [
    'prefetch:pause',
    `emit-raw:${text}`,
    `emit:${text}:tokens=yes`,
    'prefetch:resume',
  ]);
});

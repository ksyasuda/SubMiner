import test from 'node:test';
import assert from 'node:assert/strict';
import {
  captureLiveSubtitleMiningContext,
  copyCurrentSubtitle,
  handleMineSentenceDigit,
  handleMultiCopyDigit,
  mineSentenceCard,
} from './mining';
import { SubtitleTimingTracker } from '../../subtitle-timing-tracker';

test('copyCurrentSubtitle reports tracker and subtitle guards', () => {
  const osd: string[] = [];
  const copied: string[] = [];

  copyCurrentSubtitle({
    subtitleTimingTracker: null,
    writeClipboardText: (text) => copied.push(text),
    showMpvOsd: (text) => osd.push(text),
  });
  assert.equal(osd.at(-1), 'Subtitle tracker not available');

  copyCurrentSubtitle({
    subtitleTimingTracker: {
      getRecentBlocks: () => [],
      getCurrentSubtitle: () => null,
      findTiming: () => null,
    },
    writeClipboardText: (text) => copied.push(text),
    showMpvOsd: (text) => osd.push(text),
  });
  assert.equal(osd.at(-1), 'No current subtitle');
  assert.deepEqual(copied, []);
});

test('copyCurrentSubtitle copies current subtitle text', () => {
  const osd: string[] = [];
  const copied: string[] = [];

  copyCurrentSubtitle({
    subtitleTimingTracker: {
      getRecentBlocks: () => [],
      getCurrentSubtitle: () => 'hello world',
      findTiming: () => null,
    },
    writeClipboardText: (text) => copied.push(text),
    showMpvOsd: (text) => osd.push(text),
  });

  assert.deepEqual(copied, ['hello world']);
  assert.equal(osd.at(-1), 'Copied subtitle');
});

test('mineSentenceCard handles missing integration and disconnected mpv', async () => {
  const osd: string[] = [];

  assert.equal(
    await mineSentenceCard({
      ankiIntegration: null,
      mpvClient: null,
      showMpvOsd: (text) => osd.push(text),
    }),
    false,
  );
  assert.equal(osd.at(-1), 'AnkiConnect integration not enabled');

  assert.equal(
    await mineSentenceCard({
      ankiIntegration: {
        updateLastAddedFromClipboard: async () => {},
        triggerFieldGroupingForLastAddedCard: async () => {},
        markLastCardAsAudioCard: async () => {},
        createSentenceCard: async () => false,
      },
      mpvClient: {
        connected: false,
        currentSubText: 'line',
        currentSubStart: 1,
        currentSubEnd: 2,
      },
      showMpvOsd: (text) => osd.push(text),
    }),
    false,
  );

  assert.equal(osd.at(-1), 'MPV not connected');
});

test('mineSentenceCard creates sentence card from mpv subtitle state', async () => {
  const created: Array<{
    sentence: string;
    startTime: number;
    endTime: number;
    secondarySub?: string;
  }> = [];

  const createdCard = await mineSentenceCard({
    ankiIntegration: {
      updateLastAddedFromClipboard: async () => {},
      triggerFieldGroupingForLastAddedCard: async () => {},
      markLastCardAsAudioCard: async () => {},
      createSentenceCard: async (sentence, startTime, endTime, secondarySub) => {
        created.push({ sentence, startTime, endTime, secondarySub });
        return true;
      },
    },
    mpvClient: {
      connected: true,
      currentSubText: 'subtitle line',
      currentSubStart: 10,
      currentSubEnd: 12,
      currentSecondarySubText: 'secondary line',
    },
    showMpvOsd: () => {},
  });

  assert.equal(createdCard, true);
  assert.deepEqual(created, [
    {
      sentence: 'subtitle line',
      startTime: 10,
      endTime: 12,
      secondarySub: 'secondary line',
    },
  ]);
});

test('mineSentenceCard prefers a canonical primary subtitle snapshot', async () => {
  const created: Array<{
    sentence: string;
    startTime: number;
    endTime: number;
    secondarySub?: string;
  }> = [];

  await mineSentenceCard({
    ankiIntegration: {
      updateLastAddedFromClipboard: async () => {},
      triggerFieldGroupingForLastAddedCard: async () => {},
      markLastCardAsAudioCard: async () => {},
      createSentenceCard: async (sentence, startTime, endTime, secondarySub) => {
        created.push({ sentence, startTime, endTime, secondarySub });
        return true;
      },
    },
    mpvClient: {
      connected: true,
      currentSubText: '今今今手手手',
      currentSubStart: 11.4,
      currentSubEnd: 11.8,
      currentSecondarySubText: 'English subtitle',
    },
    primarySubtitle: {
      text: '今　手にある物差しでは',
      startTime: 11.13,
      endTime: 13.83,
    },
    showMpvOsd: () => {},
  });

  assert.deepEqual(created, [
    {
      sentence: '今　手にある物差しでは',
      startTime: 11.13,
      endTime: 13.83,
      secondarySub: 'English subtitle',
    },
  ]);
});

test('mineSentenceCard uses normalized secondary subtitle state instead of raw mpv text', async () => {
  const created: Array<{ sentence: string; secondarySub?: string }> = [];
  let requestedRawSecondaryText = false;

  await mineSentenceCard({
    ankiIntegration: {
      updateLastAddedFromClipboard: async () => {},
      triggerFieldGroupingForLastAddedCard: async () => {},
      markLastCardAsAudioCard: async () => {},
      createSentenceCard: async (sentence, _startTime, _endTime, secondarySub) => {
        created.push({ sentence, secondarySub });
        return true;
      },
    },
    mpvClient: {
      connected: true,
      currentSubText: '日本語字幕',
      currentSubStart: 10,
      currentSubEnd: 12,
      currentSecondarySubText: 'Your\nmosaic',
      requestProperty: async () => {
        requestedRawSecondaryText = true;
        return 'Your\nYour\nYour\nYour\nmosaic';
      },
    },
    showMpvOsd: () => {},
  });

  assert.equal(requestedRawSecondaryText, false);
  assert.deepEqual(created, [{ sentence: '日本語字幕', secondarySub: 'Your\nmosaic' }]);
});

test('mineSentenceCard omits normalized secondary text that matches the primary subtitle', async () => {
  const created: Array<{ sentence: string; secondarySub?: string }> = [];

  await mineSentenceCard({
    ankiIntegration: {
      updateLastAddedFromClipboard: async () => {},
      triggerFieldGroupingForLastAddedCard: async () => {},
      markLastCardAsAudioCard: async () => {},
      createSentenceCard: async (sentence, _startTime, _endTime, secondarySub) => {
        created.push({ sentence, secondarySub });
        return true;
      },
    },
    mpvClient: {
      connected: true,
      currentSubText: '日本語字幕',
      currentSubStart: 10,
      currentSubEnd: 12,
      currentSecondarySubText: '日本語字幕',
    },
    showMpvOsd: () => {},
  });

  assert.deepEqual(created, [{ sentence: '日本語字幕', secondarySub: undefined }]);
});

test('handleMultiCopyDigit copies available history and reports truncation', () => {
  const osd: string[] = [];
  const copied: string[] = [];

  handleMultiCopyDigit(5, {
    subtitleTimingTracker: {
      getRecentBlocks: (count) => ['a', 'b'].slice(0, count),
      getCurrentSubtitle: () => null,
      findTiming: () => null,
    },
    writeClipboardText: (text) => copied.push(text),
    showMpvOsd: (text) => osd.push(text),
  });

  assert.deepEqual(copied, ['a\n\nb']);
  assert.equal(osd.at(-1), 'Only 2 lines available, copied 2');
});

test('handleMultiCopyDigit copies backward from the current subtitle after a backward seek', () => {
  const copied: string[] = [];
  const tracker = new SubtitleTimingTracker();

  try {
    tracker.recordSubtitle('A', 1, 2);
    tracker.recordSubtitle('B', 3, 4);
    tracker.recordSubtitle('C', 5, 6);
    tracker.recordSubtitle('B', 3, 4);

    const deps = {
      subtitleTimingTracker: tracker,
      writeClipboardText: (text: string) => copied.push(text),
      showMpvOsd: () => {},
    };

    handleMultiCopyDigit(1, deps);
    handleMultiCopyDigit(2, deps);

    assert.deepEqual(copied, ['B', 'A\n\nB']);
    assert.deepEqual(tracker.getRecentEntries(2), [
      { displayText: 'A', startTime: 1, endTime: 2, secondaryText: undefined },
      { displayText: 'B', startTime: 3, endTime: 4, secondaryText: undefined },
    ]);
  } finally {
    tracker.destroy();
  }
});

test('handleMineSentenceDigit reports async create failures', async () => {
  const osd: string[] = [];
  const logs: Array<{ message: string; err: unknown }> = [];
  let cardsMined = 0;

  handleMineSentenceDigit(2, {
    subtitleTimingTracker: {
      getRecentBlocks: () => ['one', 'two'],
      getCurrentSubtitle: () => null,
      findTiming: (text) =>
        text === 'one' ? { startTime: 1, endTime: 3 } : { startTime: 4, endTime: 7 },
    },
    ankiIntegration: {
      updateLastAddedFromClipboard: async () => {},
      triggerFieldGroupingForLastAddedCard: async () => {},
      markLastCardAsAudioCard: async () => {},
      createSentenceCard: async () => {
        throw new Error('mine boom');
      },
    },
    getCurrentSecondarySubText: () => 'sub2',
    showMpvOsd: (text) => osd.push(text),
    logError: (message, err) => logs.push({ message, err }),
    onCardsMined: (count) => {
      cardsMined += count;
    },
  });

  await new Promise((resolve) => setImmediate(resolve));

  assert.equal(logs.length, 1);
  assert.equal(logs[0]?.message, 'mineSentenceMultiple failed:');
  assert.equal((logs[0]?.err as Error).message, 'mine boom');
  assert.ok(osd.some((entry) => entry.includes('Mine sentence failed: mine boom')));
  assert.equal(cardsMined, 0);
});

test('handleMineSentenceDigit increments successful card count', async () => {
  const osd: string[] = [];
  let cardsMined = 0;

  handleMineSentenceDigit(2, {
    subtitleTimingTracker: {
      getRecentBlocks: () => ['one', 'two'],
      getCurrentSubtitle: () => null,
      findTiming: (text) =>
        text === 'one' ? { startTime: 1, endTime: 3 } : { startTime: 4, endTime: 7 },
    },
    ankiIntegration: {
      updateLastAddedFromClipboard: async () => {},
      triggerFieldGroupingForLastAddedCard: async () => {},
      markLastCardAsAudioCard: async () => {},
      createSentenceCard: async () => true,
    },
    getCurrentSecondarySubText: () => 'sub2',
    showMpvOsd: (text) => osd.push(text),
    logError: () => {},
    onCardsMined: (count) => {
      cardsMined += count;
    },
  });

  await new Promise((resolve) => setImmediate(resolve));

  assert.equal(cardsMined, 1);
});

test('handleMineSentenceDigit keeps per-entry timings when subtitle text repeats', async () => {
  const created: Array<{ sentence: string; startTime: number; endTime: number }> = [];
  const tracker = new SubtitleTimingTracker();

  try {
    tracker.recordSubtitle('same', 1, 2);
    tracker.recordSubtitle('other', 3, 4);
    tracker.recordSubtitle('same', 5, 6);

    handleMineSentenceDigit(3, {
      subtitleTimingTracker: tracker,
      ankiIntegration: {
        updateLastAddedFromClipboard: async () => {},
        triggerFieldGroupingForLastAddedCard: async () => {},
        markLastCardAsAudioCard: async () => {},
        createSentenceCard: async (sentence, startTime, endTime) => {
          created.push({ sentence, startTime, endTime });
          return true;
        },
      },
      getCurrentSecondarySubText: () => undefined,
      showMpvOsd: () => {},
      logError: () => {},
    });

    await new Promise((resolve) => setImmediate(resolve));

    assert.deepEqual(created, [{ sentence: 'same other same', startTime: 1, endTime: 6 }]);
  } finally {
    tracker.destroy();
  }
});

test('subtitle timing history preserves adjacent repeated text with distinct timings', () => {
  const tracker = new SubtitleTimingTracker();

  try {
    tracker.recordSubtitle('same', 1, 2);
    tracker.recordSubtitle('same', 3, 4);

    assert.deepEqual(tracker.getRecentEntries(2), [
      { displayText: 'same', startTime: 1, endTime: 2, secondaryText: undefined },
      { displayText: 'same', startTime: 3, endTime: 4, secondaryText: undefined },
    ]);
  } finally {
    tracker.destroy();
  }
});

test('handleMineSentenceDigit joins per-entry secondary subtitles when available', async () => {
  const created: Array<{ sentence: string; secondarySub?: string }> = [];
  const tracker = new SubtitleTimingTracker();
  const recordSubtitleWithSecondary = tracker.recordSubtitle as (
    text: string,
    startTime: number,
    endTime: number,
    secondaryText?: string,
  ) => void;

  try {
    recordSubtitleWithSecondary.call(tracker, 'one', 1, 2, 'translation one');
    recordSubtitleWithSecondary.call(tracker, 'two', 3, 4, 'translation two');

    handleMineSentenceDigit(2, {
      subtitleTimingTracker: tracker,
      ankiIntegration: {
        updateLastAddedFromClipboard: async () => {},
        triggerFieldGroupingForLastAddedCard: async () => {},
        markLastCardAsAudioCard: async () => {},
        createSentenceCard: async (sentence, _startTime, _endTime, secondarySub) => {
          created.push({ sentence, secondarySub });
          return true;
        },
      },
      getCurrentSecondarySubText: () => 'current translation only',
      showMpvOsd: () => {},
      logError: () => {},
    });

    await new Promise((resolve) => setImmediate(resolve));

    assert.deepEqual(created, [
      { sentence: 'one two', secondarySub: 'translation one translation two' },
    ]);
  } finally {
    tracker.destroy();
  }
});

test('captureLiveSubtitleMiningContext snapshots the current line and timings', () => {
  const context = captureLiveSubtitleMiningContext(
    { currentSubText: ' 食べる ', currentSubStart: 12.5, currentSubEnd: 15.25 },
    () => 1234,
  );

  assert.deepEqual(context, {
    source: 'overlay',
    text: '食べる',
    startTime: 12.5,
    endTime: 15.25,
    capturedAtMs: 1234,
  });
});

test('captureLiveSubtitleMiningContext rejects missing client, empty text and bad timings', () => {
  assert.equal(captureLiveSubtitleMiningContext(null), null);
  assert.equal(
    captureLiveSubtitleMiningContext({
      currentSubText: '  ',
      currentSubStart: 1,
      currentSubEnd: 2,
    }),
    null,
  );
  assert.equal(
    captureLiveSubtitleMiningContext({
      currentSubText: 'line',
      currentSubStart: 5,
      currentSubEnd: 5,
    }),
    null,
  );
  assert.equal(
    captureLiveSubtitleMiningContext({
      currentSubText: 'line',
      currentSubStart: NaN,
      currentSubEnd: 2,
    }),
    null,
  );
});

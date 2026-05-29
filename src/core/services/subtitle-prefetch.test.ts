import assert from 'node:assert/strict';
import test from 'node:test';
import { computePriorityWindow, createSubtitlePrefetchService } from './subtitle-prefetch';
import type { SubtitleData } from '../../types';
import type { SubtitleCue } from '../../types';

function makeCues(count: number, startOffset = 0): SubtitleCue[] {
  return Array.from({ length: count }, (_, i) => ({
    startTime: startOffset + i * 5,
    endTime: startOffset + i * 5 + 4,
    text: `line-${i}`,
  }));
}

test('computePriorityWindow returns next N cues from current position', () => {
  const cues = makeCues(20);
  const window = computePriorityWindow(cues, 12.0, 5);

  assert.equal(window.length, 5);
  // Position 12.0 falls during cue 2, so the active cue should be warmed first.
  assert.equal(window[0]!.text, 'line-2');
  assert.equal(window[4]!.text, 'line-6');
});

test('computePriorityWindow clamps to remaining cues at end of file', () => {
  const cues = makeCues(5);
  const window = computePriorityWindow(cues, 18.0, 10);

  // Position 18.0 is during cue 3 (start=15), so cue 3 and cue 4 remain.
  assert.equal(window.length, 2);
  assert.equal(window[0]!.text, 'line-3');
  assert.equal(window[1]!.text, 'line-4');
});

test('computePriorityWindow returns empty when past all cues', () => {
  const cues = makeCues(3);
  const window = computePriorityWindow(cues, 999.0, 10);
  assert.equal(window.length, 0);
});

test('computePriorityWindow at position 0 returns first N cues', () => {
  const cues = makeCues(20);
  const window = computePriorityWindow(cues, 0, 5);

  assert.equal(window.length, 5);
  assert.equal(window[0]!.text, 'line-0');
});

test('computePriorityWindow includes the active cue when current position is mid-line', () => {
  const cues = makeCues(20);
  const window = computePriorityWindow(cues, 18.0, 3);

  assert.equal(window.length, 3);
  assert.equal(window[0]!.text, 'line-3');
  assert.equal(window[1]!.text, 'line-4');
  assert.equal(window[2]!.text, 'line-5');
});

function flushMicrotasks(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

test('prefetch service tokenizes priority window cues and caches them', async () => {
  const cues = makeCues(20);
  const cached: Map<string, SubtitleData> = new Map();
  let tokenizeCalls = 0;

  const service = createSubtitlePrefetchService({
    cues,
    tokenizeSubtitle: async (text) => {
      tokenizeCalls += 1;
      return { text, tokens: [] };
    },
    preCacheTokenization: (text, data) => {
      cached.set(text, data);
    },
    isCacheFull: () => false,
    priorityWindowSize: 3,
  });

  service.start(0);
  // Allow all async tokenization to complete
  for (let i = 0; i < 25; i += 1) {
    await flushMicrotasks();
  }
  service.stop();

  // Priority window (first 3) should be cached
  assert.ok(cached.has('line-0'));
  assert.ok(cached.has('line-1'));
  assert.ok(cached.has('line-2'));
});

test('prefetch service stops when cache is full', async () => {
  const cues = makeCues(20);
  let tokenizeCalls = 0;
  let cacheSize = 0;

  const service = createSubtitlePrefetchService({
    cues,
    tokenizeSubtitle: async (text) => {
      tokenizeCalls += 1;
      return { text, tokens: [] };
    },
    preCacheTokenization: () => {
      cacheSize += 1;
    },
    isCacheFull: () => cacheSize >= 5,
    priorityWindowSize: 3,
  });

  service.start(0);
  for (let i = 0; i < 30; i += 1) {
    await flushMicrotasks();
  }
  service.stop();

  // Should have stopped at 5 (cache full), not tokenized all 20
  assert.ok(tokenizeCalls <= 6, `Expected <= 6 tokenize calls, got ${tokenizeCalls}`);
});

test('prefetch service can be stopped mid-flight', async () => {
  const cues = makeCues(100);
  let tokenizeCalls = 0;

  const service = createSubtitlePrefetchService({
    cues,
    tokenizeSubtitle: async (text) => {
      tokenizeCalls += 1;
      return { text, tokens: [] };
    },
    preCacheTokenization: () => {},
    isCacheFull: () => false,
    priorityWindowSize: 3,
  });

  service.start(0);
  await flushMicrotasks();
  await flushMicrotasks();
  service.stop();
  const callsAtStop = tokenizeCalls;

  // Wait more to confirm no further calls
  for (let i = 0; i < 10; i += 1) {
    await flushMicrotasks();
  }

  assert.equal(tokenizeCalls, callsAtStop, 'No further tokenize calls after stop');
  assert.ok(tokenizeCalls < 100, 'Should not have tokenized all cues');
});

test('prefetch service onSeek re-prioritizes from new position', async () => {
  const cues = makeCues(20);
  const cachedTexts: string[] = [];

  const service = createSubtitlePrefetchService({
    cues,
    tokenizeSubtitle: async (text) => ({ text, tokens: [] }),
    preCacheTokenization: (text) => {
      cachedTexts.push(text);
    },
    isCacheFull: () => false,
    priorityWindowSize: 3,
  });

  service.start(0);
  // Let a few cues process
  for (let i = 0; i < 5; i += 1) {
    await flushMicrotasks();
  }

  // Seek to near the end
  service.onSeek(80.0);
  for (let i = 0; i < 30; i += 1) {
    await flushMicrotasks();
  }
  service.stop();

  // After seek to 80.0, cues starting after 80.0 (line-17, line-18, line-19) should appear in cached
  const hasPostSeekCue = cachedTexts.some(
    (t) => t === 'line-17' || t === 'line-18' || t === 'line-19',
  );
  assert.ok(hasPostSeekCue, 'Should have cached cues after seek position');
});

test('prefetch service still warms the priority window when cache is full', async () => {
  const cues = makeCues(20);
  const cachedTexts: string[] = [];

  const service = createSubtitlePrefetchService({
    cues,
    tokenizeSubtitle: async (text) => ({ text, tokens: [] }),
    preCacheTokenization: (text) => {
      cachedTexts.push(text);
    },
    isCacheFull: () => true,
    priorityWindowSize: 3,
  });

  service.start(0);
  for (let i = 0; i < 10; i += 1) {
    await flushMicrotasks();
  }
  service.stop();

  assert.deepEqual(cachedTexts.slice(0, 3), ['line-0', 'line-1', 'line-2']);
});

test('prefetch service pause/resume halts and continues tokenization', async () => {
  const cues = makeCues(20);
  let tokenizeCalls = 0;

  const service = createSubtitlePrefetchService({
    cues,
    tokenizeSubtitle: async (text) => {
      tokenizeCalls += 1;
      return { text, tokens: [] };
    },
    preCacheTokenization: () => {},
    isCacheFull: () => false,
    priorityWindowSize: 3,
  });

  service.start(0);
  await flushMicrotasks();
  await flushMicrotasks();
  service.pause();

  const callsWhenPaused = tokenizeCalls;
  // Wait while paused
  for (let i = 0; i < 5; i += 1) {
    await flushMicrotasks();
  }
  // Should not have advanced much (may have 1 in-flight)
  assert.ok(tokenizeCalls <= callsWhenPaused + 1, 'Should not tokenize much while paused');

  service.resume();
  for (let i = 0; i < 30; i += 1) {
    await flushMicrotasks();
  }
  service.stop();

  assert.ok(tokenizeCalls > callsWhenPaused + 1, 'Should resume tokenizing after unpause');
});

test('prefetch service skips cues already present in tokenization cache', async () => {
  const cues = makeCues(5);
  const tokenizedTexts: string[] = [];

  const service = createSubtitlePrefetchService({
    cues,
    tokenizeSubtitle: async (text) => {
      tokenizedTexts.push(text);
      return { text, tokens: [] };
    },
    preCacheTokenization: () => {},
    hasCachedTokenization: (text) => text === 'line-0' || text === 'line-1',
    isCacheFull: () => false,
    priorityWindowSize: 3,
  });

  service.start(0);
  for (let i = 0; i < 10; i += 1) {
    await flushMicrotasks();
  }
  service.stop();

  assert.ok(!tokenizedTexts.includes('line-0'));
  assert.ok(!tokenizedTexts.includes('line-1'));
  assert.ok(tokenizedTexts.includes('line-2'));
});

test('prefetch service deduplicates repeated cue text within a run', async () => {
  const cues: SubtitleCue[] = [
    { startTime: 0, endTime: 1, text: 'same' },
    { startTime: 1, endTime: 2, text: 'same' },
    { startTime: 2, endTime: 3, text: 'other' },
  ];
  const tokenizedTexts: string[] = [];

  const service = createSubtitlePrefetchService({
    cues,
    tokenizeSubtitle: async (text) => {
      tokenizedTexts.push(text);
      return { text, tokens: [] };
    },
    preCacheTokenization: () => {},
    isCacheFull: () => false,
    priorityWindowSize: 3,
  });

  service.start(0);
  for (let i = 0; i < 10; i += 1) {
    await flushMicrotasks();
  }
  service.stop();

  assert.deepEqual(
    tokenizedTexts.filter((text) => text === 'same'),
    ['same'],
  );
  assert.ok(tokenizedTexts.includes('other'));
});

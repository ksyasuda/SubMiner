import test from 'node:test';
import assert from 'node:assert/strict';
import { createSubtitleProcessingController } from './subtitle-processing-controller';
import type { SubtitleData } from '../../types';

function flushMicrotasks(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

test('subtitle processing emits plain payload immediately on cache miss, then tokenized payload', async () => {
  const emitted: SubtitleData[] = [];
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => ({ text, tokens: [] }),
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.onSubtitleChange('字幕');
  assert.deepEqual(emitted, [{ text: '字幕', tokens: null }]);
  await flushMicrotasks();
  assert.deepEqual(emitted, [
    { text: '字幕', tokens: null },
    { text: '字幕', tokens: [] },
  ]);
});

test('cache invalidation during pending tokenization does not re-emit the plain payload', async () => {
  const emitted: SubtitleData[] = [];
  const resolvers: Array<(value: SubtitleData | null) => void> = [];
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) =>
      await new Promise<SubtitleData | null>((resolve) => {
        resolvers.push(() => resolve({ text, tokens: [{ value: resolvers.length } as never] }));
      }),
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.onSubtitleChange('行');
  assert.deepEqual(emitted, [{ text: '行', tokens: null }]);

  controller.invalidateTokenizationCache();
  resolvers[0]?.({ text: '行', tokens: [] });
  await flushMicrotasks();
  // Retry for the new generation is now pending; still no duplicate plain emit.
  assert.deepEqual(emitted, [{ text: '行', tokens: null }]);

  resolvers[1]?.({ text: '行', tokens: [] });
  await flushMicrotasks();
  assert.deepEqual(emitted, [
    { text: '行', tokens: null },
    { text: '行', tokens: [{ value: 2 } as never] },
  ]);
});

test('failed refresh does not downgrade an already emitted tokenized subtitle', async () => {
  const emitted: SubtitleData[] = [];
  let tokenizeCalls = 0;
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => {
      tokenizeCalls += 1;
      if (tokenizeCalls > 1) {
        throw new Error('tokenizer gone');
      }
      return { text, tokens: [] };
    },
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.onSubtitleChange('行');
  await flushMicrotasks();
  controller.invalidateTokenizationCache();
  controller.refreshCurrentSubtitle();
  await flushMicrotasks();

  assert.equal(tokenizeCalls, 2);
  assert.deepEqual(emitted, [
    { text: '行', tokens: null },
    { text: '行', tokens: [] },
  ]);
});

test('null-tokenization refresh does not downgrade an already emitted tokenized subtitle', async () => {
  const emitted: SubtitleData[] = [];
  let tokenizeCalls = 0;
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => {
      tokenizeCalls += 1;
      return tokenizeCalls > 1 ? null : { text, tokens: [] };
    },
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.onSubtitleChange('行');
  await flushMicrotasks();
  controller.invalidateTokenizationCache();
  controller.refreshCurrentSubtitle();
  await flushMicrotasks();

  assert.equal(tokenizeCalls, 2);
  assert.deepEqual(emitted, [
    { text: '行', tokens: null },
    { text: '行', tokens: [] },
  ]);
});

test('subtitle processing does not emit plain payload for cached lines', async () => {
  const emitted: SubtitleData[] = [];
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => ({ text, tokens: [] }),
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.preCacheTokenization('字幕', { text: '字幕', tokens: [] });
  controller.onSubtitleChange('字幕');
  await flushMicrotasks();
  assert.deepEqual(emitted, [{ text: '字幕', tokens: [] }]);
});

test('subtitle processing shows plain line while tokenization is still pending', async () => {
  const emitted: SubtitleData[] = [];
  let resolveTokenization: ((value: SubtitleData | null) => void) | undefined;
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) =>
      await new Promise<SubtitleData | null>((resolve) => {
        resolveTokenization = () => resolve({ text, tokens: [] });
      }),
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.onSubtitleChange('遅い行');
  await flushMicrotasks();
  assert.deepEqual(emitted, [{ text: '遅い行', tokens: null }]);

  assert.ok(resolveTokenization);
  resolveTokenization({ text: '遅い行', tokens: [] });
  await flushMicrotasks();
  assert.deepEqual(emitted, [
    { text: '遅い行', tokens: null },
    { text: '遅い行', tokens: [] },
  ]);
});

test('subtitle processing drops stale tokenization and delivers latest subtitle only once', async () => {
  const emitted: SubtitleData[] = [];
  let firstResolve: ((value: SubtitleData | null) => void) | undefined;
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => {
      if (text === 'first') {
        return await new Promise<SubtitleData | null>((resolve) => {
          firstResolve = resolve;
        });
      }
      return { text, tokens: [] };
    },
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.onSubtitleChange('first');
  controller.onSubtitleChange('second');
  assert.ok(firstResolve);
  firstResolve({ text: 'first', tokens: [] });
  await flushMicrotasks();
  await flushMicrotasks();

  assert.deepEqual(emitted, [
    { text: 'first', tokens: null },
    { text: 'second', tokens: null },
    { text: 'second', tokens: [] },
  ]);
});

test('subtitle processing skips duplicate subtitle emission', async () => {
  const emitted: SubtitleData[] = [];
  let tokenizeCalls = 0;
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => {
      tokenizeCalls += 1;
      return { text, tokens: [] };
    },
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.onSubtitleChange('same');
  await flushMicrotasks();
  controller.onSubtitleChange('same');
  await flushMicrotasks();

  assert.deepEqual(emitted, [
    { text: 'same', tokens: null },
    { text: 'same', tokens: [] },
  ]);
  assert.equal(tokenizeCalls, 1);
});

test('subtitle processing reuses cached tokenization for repeated subtitle text', async () => {
  const emitted: SubtitleData[] = [];
  let tokenizeCalls = 0;
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => {
      tokenizeCalls += 1;
      return { text, tokens: [] };
    },
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.onSubtitleChange('first');
  await flushMicrotasks();
  controller.onSubtitleChange('second');
  await flushMicrotasks();
  controller.onSubtitleChange('first');
  await flushMicrotasks();

  assert.equal(tokenizeCalls, 2);
  assert.deepEqual(emitted, [
    { text: 'first', tokens: null },
    { text: 'first', tokens: [] },
    { text: 'second', tokens: null },
    { text: 'second', tokens: [] },
    { text: 'first', tokens: [] },
  ]);
});

test('subtitle processing falls back to plain subtitle when tokenization returns null', async () => {
  const emitted: SubtitleData[] = [];
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async () => null,
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.onSubtitleChange('fallback');
  await flushMicrotasks();

  assert.deepEqual(
    emitted,
    [{ text: 'fallback', tokens: null }],
    'plain payload should not be re-emitted when tokenization yields nothing new',
  );
});

test('null tokenization is not cached and a later cue retries tokenization', async () => {
  const emitted: SubtitleData[] = [];
  const callsByText = new Map<string, number>();
  let failNext = true;
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => {
      callsByText.set(text, (callsByText.get(text) ?? 0) + 1);
      if (text === 'fallback' && failNext) {
        failNext = false;
        return null;
      }
      return { text, tokens: [] };
    },
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.onSubtitleChange('fallback');
  await flushMicrotasks();

  assert.equal(callsByText.get('fallback'), 1);
  assert.equal(
    controller.hasCachedSubtitle('fallback'),
    false,
    'plain fallback must not be cached when tokenization yields nothing',
  );
  assert.deepEqual(emitted, [{ text: 'fallback', tokens: null }]);

  controller.onSubtitleChange('other');
  await flushMicrotasks();
  controller.onSubtitleChange('fallback');
  await flushMicrotasks();

  assert.equal(callsByText.get('fallback'), 2, 'later cue should retry tokenization');
  assert.equal(controller.hasCachedSubtitle('fallback'), true);
  assert.deepEqual(emitted.at(-1), { text: 'fallback', tokens: [] });
});

test('subtitle processing ignores duplicate current subtitle refresh without cache invalidation', async () => {
  const emitted: SubtitleData[] = [];
  let tokenizeCalls = 0;
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => {
      tokenizeCalls += 1;
      return { text, tokens: [] };
    },
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.onSubtitleChange('same');
  await flushMicrotasks();
  controller.refreshCurrentSubtitle();
  await flushMicrotasks();

  assert.equal(tokenizeCalls, 1);
  assert.deepEqual(emitted, [
    { text: 'same', tokens: null },
    { text: 'same', tokens: [] },
  ]);
});

test('subtitle processing coalesces refresh requests while current subtitle is processing', async () => {
  const emitted: SubtitleData[] = [];
  let tokenizeCalls = 0;
  let resolveTokenization: ((value: SubtitleData | null) => void) | undefined;
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => {
      tokenizeCalls += 1;
      return await new Promise<SubtitleData | null>((resolve) => {
        resolveTokenization = () => resolve({ text, tokens: [] });
      });
    },
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.onSubtitleChange('same');
  controller.refreshCurrentSubtitle();
  controller.refreshCurrentSubtitle('same');
  assert.ok(resolveTokenization);
  resolveTokenization({ text: 'same', tokens: [] });
  await flushMicrotasks();
  await flushMicrotasks();

  assert.equal(tokenizeCalls, 1);
  assert.deepEqual(emitted, [
    { text: 'same', tokens: null },
    { text: 'same', tokens: [] },
  ]);
});

test('subtitle processing refresh re-tokenizes after cache invalidation', async () => {
  const emitted: SubtitleData[] = [];
  let tokenizeCalls = 0;
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => {
      tokenizeCalls += 1;
      return { text, tokens: [{ value: tokenizeCalls } as never] };
    },
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.onSubtitleChange('same');
  await flushMicrotasks();
  controller.invalidateTokenizationCache();
  controller.refreshCurrentSubtitle();
  await flushMicrotasks();

  assert.equal(tokenizeCalls, 2);
  assert.deepEqual(emitted, [
    { text: 'same', tokens: null },
    { text: 'same', tokens: [{ value: 1 } as never] },
    { text: 'same', tokens: [{ value: 2 } as never] },
  ]);
});

test('subtitle processing refresh can use explicit text override', async () => {
  const emitted: SubtitleData[] = [];
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => ({ text, tokens: [] }),
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.refreshCurrentSubtitle('initial');
  await flushMicrotasks();

  assert.deepEqual(emitted, [
    { text: 'initial', tokens: null },
    { text: 'initial', tokens: [] },
  ]);
});

test('subtitle processing cache invalidation only affects future subtitle events', async () => {
  const emitted: SubtitleData[] = [];
  const callsByText = new Map<string, number>();
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => {
      callsByText.set(text, (callsByText.get(text) ?? 0) + 1);
      return { text, tokens: [] };
    },
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.onSubtitleChange('same');
  await flushMicrotasks();
  controller.onSubtitleChange('other');
  await flushMicrotasks();
  controller.onSubtitleChange('same');
  await flushMicrotasks();

  assert.equal(callsByText.get('same'), 1);
  assert.equal(emitted.length, 5);

  controller.invalidateTokenizationCache();
  assert.equal(emitted.length, 5);

  controller.onSubtitleChange('different');
  await flushMicrotasks();
  controller.onSubtitleChange('same');
  await flushMicrotasks();

  assert.equal(callsByText.get('same'), 2);
});

test('preCacheTokenization stores entry that is returned on next subtitle change', async () => {
  const emitted: SubtitleData[] = [];
  let tokenizeCalls = 0;
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => {
      tokenizeCalls += 1;
      return { text, tokens: [] };
    },
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.preCacheTokenization('予め', { text: '予め', tokens: [] });
  controller.onSubtitleChange('予め');
  await flushMicrotasks();

  assert.equal(tokenizeCalls, 0, 'should not call tokenize when pre-cached');
  assert.deepEqual(emitted, [{ text: '予め', tokens: [] }]);
});

test('preCacheTokenization reuses normalized subtitle text across ASS linebreak variants', async () => {
  const emitted: SubtitleData[] = [];
  let tokenizeCalls = 0;
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => {
      tokenizeCalls += 1;
      return { text, tokens: [] };
    },
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.preCacheTokenization('一行目\\N二行目', { text: '一行目\n二行目', tokens: [] });
  controller.onSubtitleChange('一行目\n二行目');
  await flushMicrotasks();

  assert.equal(tokenizeCalls, 0, 'should not call tokenize when normalized text matches');
  assert.deepEqual(emitted, [{ text: '一行目\n二行目', tokens: [] }]);
});

test('consumeCachedSubtitle returns prefetched payload and prevents reprocessing same line', async () => {
  const emitted: SubtitleData[] = [];
  let tokenizeCalls = 0;
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => {
      tokenizeCalls += 1;
      return { text, tokens: [] };
    },
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.preCacheTokenization('猫\\Nです', { text: '猫\nです', tokens: [] });

  const immediate = controller.consumeCachedSubtitle('猫\nです');
  assert.deepEqual(immediate, { text: '猫\nです', tokens: [] });

  controller.onSubtitleChange('猫\nです');
  await flushMicrotasks();

  assert.equal(
    tokenizeCalls,
    0,
    'same cached subtitle should not reprocess after immediate consume',
  );
  assert.deepEqual(emitted, []);
});

test('hasCachedSubtitle checks prefetched entries without consuming them', async () => {
  const emitted: SubtitleData[] = [];
  let tokenizeCalls = 0;
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => {
      tokenizeCalls += 1;
      return { text, tokens: [] };
    },
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.preCacheTokenization('猫\\Nです', { text: '猫\nです', tokens: [] });

  assert.equal(controller.hasCachedSubtitle('猫\nです'), true);

  controller.onSubtitleChange('猫\nです');
  await flushMicrotasks();

  assert.equal(tokenizeCalls, 0);
  assert.deepEqual(emitted, [{ text: '猫\nです', tokens: [] }]);

  controller.invalidateTokenizationCache();
  assert.equal(controller.hasCachedSubtitle('猫\nです'), false);
});

test('cache keeps every entry while below the limit', () => {
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => ({ text, tokens: [] }),
    emitSubtitle: () => {},
    cacheLimit: 8,
  });

  for (let i = 0; i < 8; i += 1) {
    controller.preCacheTokenization(`line-${i}`, { text: `line-${i}`, tokens: [] });
  }

  assert.deepEqual(
    Array.from({ length: 8 }, (_, i) => controller.hasCachedSubtitle(`line-${i}`)),
    Array.from({ length: 8 }, () => true),
  );
});

test('cache evicts least recently used entries once the limit is reached', () => {
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => ({ text, tokens: [] }),
    emitSubtitle: () => {},
    cacheLimit: 3,
  });

  for (const line of ['a', 'b', 'c']) {
    controller.preCacheTokenization(line, { text: line, tokens: [] });
  }
  // Touching 'a' makes 'b' the eviction candidate.
  controller.consumeCachedSubtitle('a');
  controller.preCacheTokenization('d', { text: 'd', tokens: [] });

  assert.equal(controller.hasCachedSubtitle('b'), false);
  assert.deepEqual(
    ['a', 'c', 'd'].map((line) => controller.hasCachedSubtitle(line)),
    [true, true, true],
  );
});

test('default cache limit covers a full-length title without evicting', () => {
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => ({ text, tokens: [] }),
    emitSubtitle: () => {},
  });

  for (let i = 0; i < 2000; i += 1) {
    controller.preCacheTokenization(`line-${i}`, { text: `line-${i}`, tokens: [] });
  }

  assert.equal(controller.hasCachedSubtitle('line-0'), true);
  assert.equal(controller.hasCachedSubtitle('line-1999'), true);
});

test('onSubtitleChange reports whether processing was scheduled', async () => {
  const emitted: SubtitleData[] = [];
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => ({ text, tokens: [] }),
    emitSubtitle: (payload) => emitted.push(payload),
  });

  // New text schedules work, so an emit (and anything gated on it) will follow.
  assert.equal(controller.onSubtitleChange('字幕'), true);
  await flushMicrotasks();

  // A repeat emits nothing, so callers must not wait on an emit that is never
  // coming (subtitle prefetching would stay paused for the rest of the cue).
  const emittedCount = emitted.length;
  assert.equal(controller.onSubtitleChange('字幕'), false);
  await flushMicrotasks();
  assert.equal(emitted.length, emittedCount);
});

test('refreshCurrentSubtitle reports the empty-text emit that an in-flight run will deliver', async () => {
  const emitted: SubtitleData[] = [];
  let resolveFirst: ((value: SubtitleData | null) => void) | undefined;
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => {
      if (text === '字幕') {
        return await new Promise<SubtitleData | null>((resolve) => {
          resolveFirst = resolve;
        });
      }
      return { text, tokens: [] };
    },
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.onSubtitleChange('字幕');
  await flushMicrotasks();

  // Clearing the subtitle while tokenization is in flight: the running loop
  // picks the empty text up and emits it, so callers gated on that emit (the
  // prefetch pause) must be told one is coming.
  assert.equal(controller.refreshCurrentSubtitle(''), true);

  resolveFirst?.({ text: '字幕', tokens: [] });
  await flushMicrotasks();
  await flushMicrotasks();
  assert.deepEqual(
    emitted.map((payload) => payload.text),
    [''],
  );
});

test('refreshCurrentSubtitle reports no emit for empty text when nothing is running', async () => {
  const emitted: SubtitleData[] = [];
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => ({ text, tokens: [] }),
    emitSubtitle: (payload) => emitted.push(payload),
  });

  assert.equal(controller.refreshCurrentSubtitle(''), false);
  await flushMicrotasks();
  assert.deepEqual(emitted, []);
});

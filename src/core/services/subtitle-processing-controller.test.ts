import test from 'node:test';
import assert from 'node:assert/strict';
import { createSubtitleProcessingController } from './subtitle-processing-controller';
import type { SubtitleData } from '../../types';

function flushMicrotasks(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

test('subtitle processing emits tokenized payload when tokenization succeeds', async () => {
  const emitted: SubtitleData[] = [];
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => ({ text, tokens: [] }),
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.onSubtitleChange('字幕');
  await flushMicrotasks();
  assert.deepEqual(emitted, [{ text: '字幕', tokens: [] }]);
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

  assert.deepEqual(emitted, [{ text: 'second', tokens: [] }]);
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

  assert.equal(emitted.length, 1);
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
    { text: 'first', tokens: [] },
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

  assert.deepEqual(emitted, [{ text: 'fallback', tokens: null }]);
});

test('subtitle processing can refresh current subtitle without text change', async () => {
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

  assert.equal(tokenizeCalls, 2);
  assert.deepEqual(emitted, [
    { text: 'same', tokens: [] },
    { text: 'same', tokens: [] },
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

  assert.deepEqual(emitted, [{ text: 'initial', tokens: [] }]);
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
  assert.equal(emitted.length, 3);

  controller.invalidateTokenizationCache();
  assert.equal(emitted.length, 3);

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

test('isCacheFull returns false when cache is below limit', () => {
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => ({ text, tokens: null }),
    emitSubtitle: () => {},
  });

  assert.equal(controller.isCacheFull(), false);
});

test('isCacheFull returns true when cache reaches limit', async () => {
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => ({ text, tokens: [] }),
    emitSubtitle: () => {},
  });

  // Fill cache to the 256 limit
  for (let i = 0; i < 256; i += 1) {
    controller.preCacheTokenization(`line-${i}`, { text: `line-${i}`, tokens: [] });
  }

  assert.equal(controller.isCacheFull(), true);
});

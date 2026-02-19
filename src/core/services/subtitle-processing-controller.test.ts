import test from 'node:test';
import assert from 'node:assert/strict';
import { createSubtitleProcessingController } from './subtitle-processing-controller';
import type { SubtitleData } from '../../types';

function flushMicrotasks(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

test('subtitle processing emits plain subtitle immediately before tokenized payload', async () => {
  const emitted: SubtitleData[] = [];
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => ({ text, tokens: [] }),
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.onSubtitleChange('字幕');
  assert.deepEqual(emitted[0], { text: '字幕', tokens: null });

  await flushMicrotasks();
  assert.deepEqual(emitted[1], { text: '字幕', tokens: [] });
});

test('subtitle processing drops stale tokenization and delivers latest subtitle only', async () => {
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

test('subtitle processing skips duplicate plain subtitle emission', async () => {
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

  const plainEmits = emitted.filter((entry) => entry.tokens === null);
  assert.equal(plainEmits.length, 1);
  assert.equal(tokenizeCalls, 1);
});

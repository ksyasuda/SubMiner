import test from 'node:test';
import assert from 'node:assert/strict';
import { createPrefetchAwareSubtitleProcessingController } from './subtitle-processing-runtime';
import type { SubtitleData } from '../../types';

function flushMicrotasks(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

test('subtitle processing resumes prefetch on settle, not on its emits', async () => {
  const events: string[] = [];
  let resolveTokenization: ((value: SubtitleData | null) => void) | undefined;
  const controller = createPrefetchAwareSubtitleProcessingController({
    tokenizeSubtitle: (text) =>
      new Promise<SubtitleData | null>((resolve) => {
        resolveTokenization = () => resolve({ text, tokens: [] });
      }),
    emitSubtitlePayload: (payload, options) =>
      events.push(`emit:${payload.text}:resume=${options.resumePrefetch}`),
    resumePrefetch: () => events.push('resume'),
    logDebug: () => {},
    now: () => 0,
  });

  controller.onSubtitleChange('一行目');
  await flushMicrotasks();
  // The provisional plain payload is out, but tokenization is still pending.
  assert.deepEqual(events, ['emit:一行目:resume=false']);

  resolveTokenization?.(null);
  await flushMicrotasks();
  await flushMicrotasks();
  assert.deepEqual(events, ['emit:一行目:resume=false', 'emit:一行目:resume=false', 'resume']);
});

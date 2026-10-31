import test from 'node:test';
import assert from 'node:assert/strict';
import { createPrefetchAwareSubtitleProcessingController } from './subtitle-processing-runtime';
import type { SubtitleData } from '../../types';

function flushMicrotasks(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

function createHarness() {
  const events: string[] = [];
  let resolveTokenization: ((value: SubtitleData | null) => void) | undefined;
  const controller = createPrefetchAwareSubtitleProcessingController({
    tokenizeSubtitle: () =>
      new Promise<SubtitleData | null>((resolve) => {
        resolveTokenization = resolve;
      }),
    emitSubtitlePayload: (payload, options) =>
      events.push(
        `emit:${payload.text}:tokens=${payload.tokens ? 'yes' : 'no'}:resume=${options.resumePrefetch}`,
      ),
    resumePrefetch: () => events.push('resume'),
    logDebug: () => {},
    now: () => 0,
  });
  return {
    events,
    controller,
    resolveTokenization: (value: SubtitleData | null) => resolveTokenization?.(value),
  };
}

test('subtitle processing resumes prefetch after the tokenized emit settles', async () => {
  const { events, controller, resolveTokenization } = createHarness();

  controller.onSubtitleChange('一行目');
  await flushMicrotasks();
  // The provisional plain payload is out, but tokenization is still pending.
  assert.deepEqual(events, ['emit:一行目:tokens=no:resume=false']);

  resolveTokenization({ text: '一行目', tokens: [] });
  await flushMicrotasks();
  await flushMicrotasks();
  assert.deepEqual(events, [
    'emit:一行目:tokens=no:resume=false',
    'emit:一行目:tokens=yes:resume=false',
    'resume',
  ]);
});

test('subtitle processing resumes prefetch when tokenization fails without another emit', async () => {
  const { events, controller, resolveTokenization } = createHarness();

  controller.onSubtitleChange('一行目');
  await flushMicrotasks();
  assert.deepEqual(events, ['emit:一行目:tokens=no:resume=false']);

  resolveTokenization(null);
  await flushMicrotasks();
  await flushMicrotasks();
  assert.deepEqual(events, ['emit:一行目:tokens=no:resume=false', 'resume']);
});

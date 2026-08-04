import test from 'node:test';
import assert from 'node:assert/strict';
import type { SubtitleData } from '../../types';
import { createSubtitleProcessingController } from './subtitle-processing-controller';

function flushMicrotasks(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

test('new subtitle emits plain immediately without parallel tokenization or a stale overwrite', async () => {
  const emitted: SubtitleData[] = [];
  const resolvers = new Map<string, (value: SubtitleData | null) => void>();
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) =>
      await new Promise<SubtitleData | null>((resolve) => {
        resolvers.set(text, resolve);
      }),
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.onSubtitleChange('first');
  controller.onSubtitleChange('second');

  assert.deepEqual(emitted, [
    { text: 'first', tokens: null },
    { text: 'second', tokens: null },
  ]);
  assert.equal(resolvers.has('second'), false);

  const resolveFirst = resolvers.get('first');
  assert.ok(resolveFirst);
  resolveFirst({ text: 'first', tokens: [] });
  await flushMicrotasks();
  assert.deepEqual(emitted, [
    { text: 'first', tokens: null },
    { text: 'second', tokens: null },
  ]);
  assert.equal(resolvers.has('second'), true);

  const resolveSecond = resolvers.get('second');
  assert.ok(resolveSecond);
  resolveSecond({ text: 'second', tokens: [] });
  await flushMicrotasks();
  assert.deepEqual(emitted, [
    { text: 'first', tokens: null },
    { text: 'second', tokens: null },
    { text: 'second', tokens: [] },
  ]);
});

test('subtitle clears immediately while previous tokenization remains pending', async () => {
  const emitted: SubtitleData[] = [];
  let resolveTokenization: ((value: SubtitleData | null) => void) | undefined;
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async () =>
      await new Promise<SubtitleData | null>((resolve) => {
        resolveTokenization = resolve;
      }),
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.onSubtitleChange('first');
  controller.onSubtitleChange('');

  assert.deepEqual(emitted, [
    { text: 'first', tokens: null },
    { text: '', tokens: null },
  ]);

  assert.ok(resolveTokenization);
  resolveTokenization({ text: 'first', tokens: [] });
  await flushMicrotasks();
  assert.deepEqual(emitted, [
    { text: 'first', tokens: null },
    { text: '', tokens: null },
  ]);
});

test('returning to an uncached completed line emits it while another line is pending', async () => {
  const emitted: SubtitleData[] = [];
  let resolvePending: ((value: SubtitleData | null) => void) | undefined;
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => {
      if (text === 'A') {
        return { text, tokens: [] };
      }
      return await new Promise<SubtitleData | null>((resolve) => {
        resolvePending = resolve;
      });
    },
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.onSubtitleChange('A');
  await flushMicrotasks();
  controller.invalidateTokenizationCache();
  controller.onSubtitleChange('B');
  controller.onSubtitleChange('A');

  assert.deepEqual(emitted.at(-1), { text: 'A', tokens: null });
  assert.ok(resolvePending);
  resolvePending({ text: 'B', tokens: [] });
});

test('ABA subtitle changes reuse the matching first tokenization only after A is current again', async () => {
  const emitted: SubtitleData[] = [];
  const tokenizeCalls: string[] = [];
  const resolvers: Array<(value: SubtitleData | null) => void> = [];
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => {
      tokenizeCalls.push(text);
      return await new Promise<SubtitleData | null>((resolve) => {
        resolvers.push(resolve);
      });
    },
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.onSubtitleChange('A');
  controller.onSubtitleChange('B');
  controller.onSubtitleChange('A');
  const resolveFirst = resolvers[0];
  assert.ok(resolveFirst);
  resolveFirst({ text: 'A', tokens: [{ value: 1 } as never] });
  await flushMicrotasks();

  assert.deepEqual(tokenizeCalls, ['A']);
  assert.deepEqual(emitted, [
    { text: 'A', tokens: null },
    { text: 'B', tokens: null },
    { text: 'A', tokens: null },
    { text: 'A', tokens: [{ value: 1 } as never] },
  ]);
});

test('cached next subtitle does not downgrade to plain while processing is busy', async () => {
  const emitted: SubtitleData[] = [];
  let resolveTokenization: ((value: SubtitleData | null) => void) | undefined;
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) =>
      await new Promise<SubtitleData | null>((resolve) => {
        resolveTokenization = () => resolve({ text, tokens: [] });
      }),
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.preCacheTokenization('cached', { text: 'cached', tokens: [] });
  controller.onSubtitleChange('pending');
  controller.onSubtitleChange('cached');

  assert.deepEqual(emitted, [{ text: 'pending', tokens: null }]);
  assert.ok(resolveTokenization);
  resolveTokenization({ text: 'pending', tokens: [] });
  await flushMicrotasks();
  assert.deepEqual(emitted, [
    { text: 'pending', tokens: null },
    { text: 'cached', tokens: [] },
  ]);
});

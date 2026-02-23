import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildSubtitleProcessingControllerMainDepsHandler } from './subtitle-processing-main-deps';

test('subtitle processing main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const deps = createBuildSubtitleProcessingControllerMainDepsHandler({
    tokenizeSubtitle: async (text) => {
      calls.push(`tokenize:${text}`);
      return { text, tokens: null };
    },
    emitSubtitle: (payload) => calls.push(`emit:${payload.text}`),
    logDebug: (message) => calls.push(`log:${message}`),
    now: () => 42,
  })();

  const tokenized = await deps.tokenizeSubtitle('line');
  deps.emitSubtitle({ text: 'line', tokens: null });
  deps.logDebug?.('ok');
  assert.equal(deps.now?.(), 42);
  assert.deepEqual(tokenized, { text: 'line', tokens: null });
  assert.deepEqual(calls, ['tokenize:line', 'emit:line', 'log:ok']);
});

test('subtitle processing main deps builder preserves optional callbacks when absent', () => {
  const deps = createBuildSubtitleProcessingControllerMainDepsHandler({
    tokenizeSubtitle: async () => null,
    emitSubtitle: () => {},
  })();

  assert.equal(deps.logDebug, undefined);
  assert.equal(deps.now, undefined);
});

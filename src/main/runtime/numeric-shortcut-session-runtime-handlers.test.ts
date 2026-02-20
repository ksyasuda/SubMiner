import assert from 'node:assert/strict';
import test from 'node:test';
import { createNumericShortcutSessionRuntimeHandlers } from './numeric-shortcut-session-runtime-handlers';

test('numeric shortcut session runtime handlers compose cancel/start handlers', () => {
  const calls: string[] = [];
  const createSession = (name: string) => ({
    start: ({ timeoutMs, onDigit }: { timeoutMs: number; onDigit: (digit: number) => void }) => {
      calls.push(`${name}:start:${timeoutMs}`);
      onDigit(3);
    },
    cancel: () => calls.push(`${name}:cancel`),
  });

  const runtime = createNumericShortcutSessionRuntimeHandlers({
    multiCopySession: createSession('multi-copy'),
    mineSentenceSession: createSession('mine-sentence'),
    onMultiCopyDigit: (count) => calls.push(`multi-copy:digit:${count}`),
    onMineSentenceDigit: (count) => calls.push(`mine-sentence:digit:${count}`),
  });

  runtime.cancelPendingMultiCopy();
  runtime.startPendingMultiCopy(500);
  runtime.cancelPendingMineSentenceMultiple();
  runtime.startPendingMineSentenceMultiple(700);

  assert.deepEqual(calls, [
    'multi-copy:cancel',
    'multi-copy:start:500',
    'multi-copy:digit:3',
    'mine-sentence:cancel',
    'mine-sentence:start:700',
    'mine-sentence:digit:3',
  ]);
});

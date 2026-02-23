import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createCancelNumericShortcutSessionHandler,
  createStartNumericShortcutSessionHandler,
} from './numeric-shortcut-session-handlers';

test('cancel numeric shortcut session handler cancels active session', () => {
  const calls: string[] = [];
  const cancel = createCancelNumericShortcutSessionHandler({
    session: {
      start: () => {},
      cancel: () => calls.push('cancel'),
    },
  });

  cancel();
  assert.deepEqual(calls, ['cancel']);
});

test('start numeric shortcut session handler forwards timeout, messages, and onDigit', () => {
  const calls: string[] = [];
  const start = createStartNumericShortcutSessionHandler({
    session: {
      cancel: () => {},
      start: ({ timeoutMs, onDigit, messages }) => {
        calls.push(`timeout:${timeoutMs}`);
        calls.push(`prompt:${messages.prompt}`);
        onDigit(3);
      },
    },
    onDigit: (digit) => calls.push(`digit:${digit}`),
    messages: {
      prompt: 'Prompt',
      timeout: 'Timeout',
      cancelled: 'Cancelled',
    },
  });

  start(1200);
  assert.deepEqual(calls, ['timeout:1200', 'prompt:Prompt', 'digit:3']);
});

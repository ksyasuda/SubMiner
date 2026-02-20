import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildCancelNumericShortcutSessionMainDepsHandler,
  createBuildStartNumericShortcutSessionMainDepsHandler,
} from './numeric-shortcut-session-main-deps';

test('numeric shortcut session main deps builders map callbacks', () => {
  const calls: string[] = [];
  const session = {
    start: () => calls.push('start'),
    cancel: () => calls.push('cancel'),
  };

  const cancel = createBuildCancelNumericShortcutSessionMainDepsHandler({ session })();
  cancel.session.cancel();

  const start = createBuildStartNumericShortcutSessionMainDepsHandler({
    session,
    onDigit: (digit) => calls.push(`digit:${digit}`),
    messages: {
      prompt: 'prompt',
      timeout: 'timeout',
      cancelled: 'cancelled',
    },
  })();
  start.session.start({
    timeoutMs: 100,
    onDigit: () => {},
    messages: start.messages,
  });
  start.onDigit(4);
  assert.equal(start.messages.prompt, 'prompt');
  assert.equal(start.messages.timeout, 'timeout');
  assert.equal(start.messages.cancelled, 'cancelled');

  assert.deepEqual(calls, ['cancel', 'start', 'digit:4']);
});

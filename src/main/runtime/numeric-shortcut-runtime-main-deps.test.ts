import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildNumericShortcutRuntimeMainDepsHandler } from './numeric-shortcut-runtime-main-deps';

test('numeric shortcut runtime main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildNumericShortcutRuntimeMainDepsHandler({
    globalShortcut: {
      register: () => true,
      unregister: () => {
        calls.push('unregister');
      },
    },
    showMpvOsd: (text) => calls.push(`osd:${text}`),
    setTimer: (handler) => {
      calls.push('timer');
      handler();
      return 1 as never;
    },
    clearTimer: () => {
      calls.push('clear');
    },
  })();

  assert.equal(
    deps.globalShortcut.register('1', () => {}),
    true,
  );
  deps.globalShortcut.unregister('1');
  deps.showMpvOsd('x');
  deps.setTimer(() => calls.push('tick'), 1000);
  deps.clearTimer(1 as never);

  assert.deepEqual(calls, ['unregister', 'osd:x', 'timer', 'tick', 'clear']);
});

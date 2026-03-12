import assert from 'node:assert/strict';
import test from 'node:test';

// @ts-expect-error Vendor Yomitan modules are JS-only in this repo.
import { Display } from '../../vendor/subminer-yomitan/ext/js/display/display.js';

test('yomitan display scroll bridge uses popup scroll container instead of window scroll', () => {
  let scrolledTo: { x: number; y: number } | null = null;
  const result = Display.prototype._onMessageScrollBy.call(
    {
      _windowScroll: {
        x: 24,
        y: 80,
        to(x: number, y: number) {
          scrolledTo = { x, y };
        },
      },
    },
    { deltaX: 12, deltaY: -20 },
  );

  assert.equal(result, true);
  assert.deepEqual(scrolledTo, { x: 36, y: 60 });
});

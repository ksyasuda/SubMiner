import assert from 'node:assert/strict';
import test from 'node:test';
import { clearYomitanParserRuntimeState } from './yomitan-extension-runtime-state';

const cases = [
  { name: 'a live parser window', window: { destroyed: false }, expectDestroyed: true },
  {
    name: 'an already destroyed parser window',
    window: { destroyed: true },
    expectDestroyed: false,
  },
  { name: 'no parser window', window: null, expectDestroyed: false },
];

for (const c of cases) {
  test(`clearYomitanParserRuntimeState clears parser state with ${c.name}`, () => {
    let destroyCalls = 0;
    const parserWindow = c.window && {
      isDestroyed: () => c.window.destroyed,
      destroy: () => {
        destroyCalls += 1;
      },
    };
    const state: {
      window: unknown;
      ready: Promise<void> | null;
      init: Promise<boolean> | null;
    } = {
      window: parserWindow,
      ready: Promise.resolve(),
      init: Promise.resolve(true),
    };

    clearYomitanParserRuntimeState({
      getYomitanParserWindow: () => parserWindow,
      setYomitanParserWindow: (window) => {
        state.window = window;
      },
      setYomitanParserReadyPromise: (promise) => {
        state.ready = promise;
      },
      setYomitanParserInitPromise: (promise) => {
        state.init = promise;
      },
    });

    assert.equal(destroyCalls, c.expectDestroyed ? 1 : 0);
    assert.deepEqual(state, { window: null, ready: null, init: null });
  });
}

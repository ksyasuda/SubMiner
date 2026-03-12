import assert from 'node:assert/strict';
import test from 'node:test';
import { clearYomitanParserRuntimeState } from './yomitan-extension-runtime-state';

test('clearYomitanParserRuntimeState destroys parser window and clears parser promises', () => {
  const calls: string[] = [];
  const parserWindow = {
    isDestroyed: () => false,
    destroy: () => {
      calls.push('destroy');
    },
  };

  clearYomitanParserRuntimeState({
    getYomitanParserWindow: () => parserWindow as never,
    setYomitanParserWindow: (window) => calls.push(`window:${window === null ? 'null' : 'set'}`),
    setYomitanParserReadyPromise: (promise) =>
      calls.push(`ready:${promise === null ? 'null' : 'set'}`),
    setYomitanParserInitPromise: (promise) =>
      calls.push(`init:${promise === null ? 'null' : 'set'}`),
  });

  assert.deepEqual(calls, ['destroy', 'window:null', 'ready:null', 'init:null']);
});

test('clearYomitanParserRuntimeState skips destroy when parser window is already gone', () => {
  const calls: string[] = [];
  const parserWindow = {
    isDestroyed: () => true,
    destroy: () => {
      calls.push('destroy');
    },
  };

  clearYomitanParserRuntimeState({
    getYomitanParserWindow: () => parserWindow as never,
    setYomitanParserWindow: (window) => calls.push(`window:${window === null ? 'null' : 'set'}`),
    setYomitanParserReadyPromise: (promise) =>
      calls.push(`ready:${promise === null ? 'null' : 'set'}`),
    setYomitanParserInitPromise: (promise) =>
      calls.push(`init:${promise === null ? 'null' : 'set'}`),
  });

  assert.deepEqual(calls, ['window:null', 'ready:null', 'init:null']);
});

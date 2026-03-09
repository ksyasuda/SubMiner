import assert from 'node:assert/strict';
import test from 'node:test';
import {
  registerSecondInstanceHandlerEarly,
  requestSingleInstanceLockEarly,
  resetEarlySingleInstanceStateForTests,
} from './early-single-instance';

function createFakeApp(lockValue = true) {
  let requestCalls = 0;
  let secondInstanceListener: ((_event: unknown, argv: string[]) => void) | null = null;

  return {
    app: {
      requestSingleInstanceLock: () => {
        requestCalls += 1;
        return lockValue;
      },
      on: (_event: 'second-instance', listener: (_event: unknown, argv: string[]) => void) => {
        secondInstanceListener = listener;
      },
    },
    emitSecondInstance: (argv: string[]) => {
      secondInstanceListener?.({}, argv);
    },
    getRequestCalls: () => requestCalls,
  };
}

test('requestSingleInstanceLockEarly caches the lock result per process', () => {
  resetEarlySingleInstanceStateForTests();
  const fake = createFakeApp(true);

  assert.equal(requestSingleInstanceLockEarly(fake.app), true);
  assert.equal(requestSingleInstanceLockEarly(fake.app), true);
  assert.equal(fake.getRequestCalls(), 1);
});

test('registerSecondInstanceHandlerEarly replays queued argv and forwards new events', () => {
  resetEarlySingleInstanceStateForTests();
  const fake = createFakeApp(true);
  const calls: string[][] = [];

  assert.equal(requestSingleInstanceLockEarly(fake.app), true);
  fake.emitSecondInstance(['SubMiner.exe', '--start', '--socket', '\\\\.\\pipe\\subminer']);

  registerSecondInstanceHandlerEarly(fake.app, (_event, argv) => {
    calls.push(argv);
  });
  fake.emitSecondInstance(['SubMiner.exe', '--start', '--show-visible-overlay']);

  assert.deepEqual(calls, [
    ['SubMiner.exe', '--start', '--socket', '\\\\.\\pipe\\subminer'],
    ['SubMiner.exe', '--start', '--show-visible-overlay'],
  ]);
});

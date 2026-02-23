import test from 'node:test';
import assert from 'node:assert/strict';
import { registerProtocolUrlHandlers } from './protocol-url-handlers';

test('registerProtocolUrlHandlers wires open-url and second-instance handling', () => {
  const listeners = new Map<string, (...args: unknown[]) => void>();
  const calls: string[] = [];
  registerProtocolUrlHandlers({
    registerOpenUrl: (listener) => {
      listeners.set('open-url', listener as (...args: unknown[]) => void);
    },
    registerSecondInstance: (listener) => {
      listeners.set('second-instance', listener as (...args: unknown[]) => void);
    },
    handleAnilistSetupProtocolUrl: (rawUrl) => rawUrl.includes('anilist-setup'),
    findAnilistSetupDeepLinkArgvUrl: (argv) =>
      argv.find((entry) => entry.startsWith('subminer://')) ?? null,
    logUnhandledOpenUrl: (rawUrl) => calls.push(`open:${rawUrl}`),
    logUnhandledSecondInstanceUrl: (rawUrl) => calls.push(`second:${rawUrl}`),
  });

  const openUrlListener = listeners.get('open-url');
  const secondInstanceListener = listeners.get('second-instance');
  if (!openUrlListener || !secondInstanceListener) {
    throw new Error('expected listeners');
  }

  let prevented = false;
  openUrlListener({ preventDefault: () => (prevented = true) }, 'subminer://noop');
  secondInstanceListener({}, ['foo', 'subminer://noop']);

  assert.equal(prevented, true);
  assert.deepEqual(calls, ['open:subminer://noop', 'second:subminer://noop']);
});

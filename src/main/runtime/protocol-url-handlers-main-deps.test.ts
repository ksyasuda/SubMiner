import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildRegisterProtocolUrlHandlersMainDepsHandler } from './protocol-url-handlers-main-deps';

test('protocol url handlers main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildRegisterProtocolUrlHandlersMainDepsHandler({
    registerOpenUrl: () => calls.push('open-register'),
    registerSecondInstance: () => calls.push('second-register'),
    handleAnilistSetupProtocolUrl: () => true,
    findAnilistSetupDeepLinkArgvUrl: () => 'subminer://anilist-setup',
    logUnhandledOpenUrl: (rawUrl) => calls.push(`open:${rawUrl}`),
    logUnhandledSecondInstanceUrl: (rawUrl) => calls.push(`second:${rawUrl}`),
  })();

  deps.registerOpenUrl(() => {});
  deps.registerSecondInstance(() => {});
  assert.equal(deps.handleAnilistSetupProtocolUrl('subminer://anilist-setup'), true);
  assert.equal(deps.findAnilistSetupDeepLinkArgvUrl(['x']), 'subminer://anilist-setup');
  deps.logUnhandledOpenUrl('subminer://noop');
  deps.logUnhandledSecondInstanceUrl('subminer://noop');

  assert.deepEqual(calls, [
    'open-register',
    'second-register',
    'open:subminer://noop',
    'second:subminer://noop',
  ]);
});

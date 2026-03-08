import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildConsumeAnilistSetupTokenFromUrlMainDepsHandler,
  createBuildHandleAnilistSetupProtocolUrlMainDepsHandler,
  createBuildNotifyAnilistSetupMainDepsHandler,
  createBuildRegisterSubminerProtocolClientMainDepsHandler,
} from './anilist-setup-protocol-main-deps';

test('notify anilist setup main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildNotifyAnilistSetupMainDepsHandler({
    hasMpvClient: () => true,
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    showDesktopNotification: (title) => calls.push(`notify:${title}`),
    logInfo: (message) => calls.push(`log:${message}`),
  })();

  assert.equal(deps.hasMpvClient(), true);
  deps.showMpvOsd('ok');
  deps.showDesktopNotification('SubMiner', { body: 'x' });
  deps.logInfo('done');
  assert.deepEqual(calls, ['osd:ok', 'notify:SubMiner', 'log:done']);
});

test('consume anilist setup token main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildConsumeAnilistSetupTokenFromUrlMainDepsHandler({
    consumeAnilistSetupCallbackUrl: () => true,
    saveToken: () => calls.push('save'),
    setCachedToken: () => calls.push('cache'),
    setResolvedState: () => calls.push('resolved'),
    setSetupPageOpened: () => calls.push('opened'),
    onSuccess: () => calls.push('success'),
    closeWindow: () => calls.push('close'),
  })();

  assert.equal(
    deps.consumeAnilistSetupCallbackUrl({
      rawUrl: 'subminer://anilist-setup',
      saveToken: () => {},
      setCachedToken: () => {},
      setResolvedState: () => {},
      setSetupPageOpened: () => {},
      onSuccess: () => {},
      closeWindow: () => {},
    }),
    true,
  );
  deps.saveToken('token');
  deps.setCachedToken('token');
  deps.setResolvedState(Date.now());
  deps.setSetupPageOpened(true);
  deps.onSuccess();
  deps.closeWindow();
  assert.deepEqual(calls, ['save', 'cache', 'resolved', 'opened', 'success', 'close']);
});

test('handle anilist setup protocol url main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildHandleAnilistSetupProtocolUrlMainDepsHandler({
    consumeAnilistSetupTokenFromUrl: () => true,
    logWarn: (message) => calls.push(`warn:${message}`),
  })();

  assert.equal(deps.consumeAnilistSetupTokenFromUrl('subminer://anilist-setup'), true);
  deps.logWarn('missing', null);
  assert.deepEqual(calls, ['warn:missing']);
});

test('register subminer protocol client main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildRegisterSubminerProtocolClientMainDepsHandler({
    isDefaultApp: () => true,
    getArgv: () => ['electron', 'entry.js'],
    execPath: '/tmp/electron',
    resolvePath: (value) => `/abs/${value}`,
    setAsDefaultProtocolClient: () => true,
    logDebug: (message) => calls.push(`debug:${message}`),
  })();

  assert.equal(deps.isDefaultApp(), true);
  assert.deepEqual(deps.getArgv(), ['electron', 'entry.js']);
  assert.equal(deps.execPath, '/tmp/electron');
  assert.equal(deps.resolvePath('entry.js'), '/abs/entry.js');
  assert.equal(deps.setAsDefaultProtocolClient('subminer'), true);
});

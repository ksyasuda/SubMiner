import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildRefreshAnilistClientSecretStateMainDepsHandler } from './anilist-token-refresh-main-deps';

test('refresh anilist client secret state main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const config = { anilist: { accessToken: 'token' } };
  const deps = createBuildRefreshAnilistClientSecretStateMainDepsHandler({
    getResolvedConfig: () => config as never,
    isAnilistTrackingEnabled: () => true,
    getCachedAccessToken: () => 'cached',
    setCachedAccessToken: () => calls.push('set-cache'),
    saveStoredToken: () => calls.push('save'),
    loadStoredToken: () => 'stored',
    setClientSecretState: () => calls.push('set-state'),
    getAnilistSetupPageOpened: () => false,
    setAnilistSetupPageOpened: () => calls.push('set-opened'),
    openAnilistSetupWindow: () => calls.push('open-window'),
    now: () => 123,
  })();

  assert.equal(deps.getResolvedConfig(), config);
  assert.equal(deps.isAnilistTrackingEnabled(config as never), true);
  assert.equal(deps.getCachedAccessToken(), 'cached');
  deps.setCachedAccessToken(null);
  deps.saveStoredToken('x');
  assert.equal(deps.loadStoredToken(), 'stored');
  deps.setClientSecretState({} as never);
  assert.equal(deps.getAnilistSetupPageOpened(), false);
  deps.setAnilistSetupPageOpened(true);
  deps.openAnilistSetupWindow();
  assert.equal(deps.now(), 123);
  assert.deepEqual(calls, ['set-cache', 'save', 'set-state', 'set-opened', 'open-window']);
});

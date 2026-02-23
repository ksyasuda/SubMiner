import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildOpenAnilistSetupWindowMainDepsHandler } from './anilist-setup-window-main-deps';

test('open anilist setup window main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildOpenAnilistSetupWindowMainDepsHandler({
    maybeFocusExistingSetupWindow: () => false,
    createSetupWindow: () => ({}) as never,
    buildAuthorizeUrl: () => 'https://anilist.co/auth',
    consumeCallbackUrl: () => true,
    openSetupInBrowser: (url) => calls.push(`browser:${url}`),
    loadManualTokenEntry: () => calls.push('manual'),
    redirectUri: 'subminer://anilist-auth',
    developerSettingsUrl: 'https://anilist.co/settings/developer',
    isAllowedExternalUrl: () => true,
    isAllowedNavigationUrl: () => true,
    logWarn: (message) => calls.push(`warn:${message}`),
    logError: (message) => calls.push(`error:${message}`),
    clearSetupWindow: () => calls.push('clear'),
    setSetupPageOpened: (opened) => calls.push(`opened:${String(opened)}`),
    setSetupWindow: () => calls.push('window'),
    openExternal: (url) => calls.push(`external:${url}`),
  })();

  assert.equal(deps.maybeFocusExistingSetupWindow(), false);
  assert.equal(deps.buildAuthorizeUrl(), 'https://anilist.co/auth');
  assert.equal(deps.consumeCallbackUrl('subminer://anilist-setup?access_token=x'), true);
  assert.equal(deps.redirectUri, 'subminer://anilist-auth');
  assert.equal(deps.developerSettingsUrl, 'https://anilist.co/settings/developer');
  assert.equal(deps.isAllowedExternalUrl('https://anilist.co'), true);
  assert.equal(deps.isAllowedNavigationUrl('https://anilist.co/oauth'), true);
  deps.openSetupInBrowser('https://anilist.co/auth');
  deps.loadManualTokenEntry({} as never, 'https://anilist.co/auth');
  deps.logWarn('warn');
  deps.logError('error', null);
  deps.clearSetupWindow();
  deps.setSetupPageOpened(true);
  deps.setSetupWindow({} as never);
  deps.openExternal('https://anilist.co');

  assert.deepEqual(calls, [
    'browser:https://anilist.co/auth',
    'manual',
    'warn:warn',
    'error:error',
    'clear',
    'opened:true',
    'window',
    'external:https://anilist.co',
  ]);
});

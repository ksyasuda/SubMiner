import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildOpenJellyfinSetupWindowMainDepsHandler } from './jellyfin-setup-window-main-deps';

test('open jellyfin setup window main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const deps = createBuildOpenJellyfinSetupWindowMainDepsHandler({
    maybeFocusExistingSetupWindow: () => false,
    createSetupWindow: () => ({}) as never,
    getResolvedJellyfinConfig: () => ({ serverUrl: 'http://127.0.0.1:8096', username: 'alice' }),
    buildSetupFormHtml: () => '<html></html>',
    parseSubmissionUrl: () => ({ action: 'login', server: 's', username: 'u', password: 'p' }),
    authenticateWithPassword: async () => ({
      serverUrl: 'http://127.0.0.1:8096',
      username: 'alice',
      accessToken: 'token',
      userId: 'uid',
    }),
    getJellyfinClientInfo: () => ({
      clientName: 'SubMiner',
      clientVersion: '1.0',
      deviceId: 'dev',
    }),
    saveStoredSession: () => calls.push('save'),
    clearStoredSession: () => calls.push('clear-session'),
    patchJellyfinConfig: () => calls.push('patch'),
    logInfo: (message) => calls.push(`info:${message}`),
    logError: (message) => calls.push(`error:${message}`),
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    clearSetupWindow: () => calls.push('clear'),
    setSetupWindow: () => calls.push('set-window'),
    encodeURIComponent: (value) => encodeURIComponent(value),
    defaultServerUrl: 'http://127.0.0.1:8096',
    hasStoredSession: () => true,
  })();

  assert.equal(deps.maybeFocusExistingSetupWindow(), false);
  assert.deepEqual(deps.getResolvedJellyfinConfig(), {
    serverUrl: 'http://127.0.0.1:8096',
    username: 'alice',
  });
  assert.equal(
    deps.buildSetupFormHtml({
      servers: [],
      selectedServerUrl: 'a',
      username: 'b',
      hasStoredSession: false,
      statusMessage: '',
      statusKind: 'idle',
    }),
    '<html></html>',
  );
  assert.deepEqual(deps.parseSubmissionUrl('subminer://jellyfin-setup?x=1'), {
    action: 'login',
    server: 's',
    username: 'u',
    password: 'p',
  });
  assert.deepEqual(
    await deps.authenticateWithPassword('s', 'u', 'p', deps.getJellyfinClientInfo()),
    {
      serverUrl: 'http://127.0.0.1:8096',
      username: 'alice',
      accessToken: 'token',
      userId: 'uid',
    },
  );
  deps.saveStoredSession({ accessToken: 'token', userId: 'uid' });
  deps.clearStoredSession();
  deps.patchJellyfinConfig({
    serverUrl: 'http://127.0.0.1:8096',
    username: 'alice',
    accessToken: 'token',
    userId: 'uid',
  });
  deps.logInfo('ok');
  deps.logError('bad', null);
  deps.showMpvOsd('toast');
  deps.clearSetupWindow();
  deps.setSetupWindow({} as never);
  assert.equal(deps.encodeURIComponent('a b'), 'a%20b');
  assert.equal(deps.defaultServerUrl, 'http://127.0.0.1:8096');
  assert.equal(deps.hasStoredSession(), true);
  assert.deepEqual(calls, [
    'save',
    'clear-session',
    'patch',
    'info:ok',
    'error:bad',
    'osd:toast',
    'clear',
    'set-window',
  ]);
});

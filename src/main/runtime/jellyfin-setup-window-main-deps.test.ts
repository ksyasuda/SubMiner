import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildOpenJellyfinSetupWindowMainDepsHandler } from './jellyfin-setup-window-main-deps';

test('open jellyfin setup window main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const expectedState = {
    selectedServerUrl: 'a',
    username: 'b',
    hasStoredSession: false,
    statusMessage: '',
    statusKind: 'idle' as const,
  };
  let capturedBuildState: unknown = null;
  let capturedParseUrl = '';
  const deps = createBuildOpenJellyfinSetupWindowMainDepsHandler({
    maybeFocusExistingSetupWindow: () => false,
    createSetupWindow: () => ({}) as never,
    getResolvedJellyfinConfig: () => ({ serverUrl: 'http://127.0.0.1:8096', username: 'alice' }),
    buildSetupFormHtml: (state) => {
      capturedBuildState = state;
      return '<html></html>';
    },
    parseSubmissionUrl: (rawUrl) => {
      capturedParseUrl = rawUrl;
      return { action: 'login', server: 's', username: 'u', password: 'p' };
    },
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
    persistAuthenticatedSession: () => calls.push('persist'),
    restartRemoteSession: () => {
      calls.push('restart-remote');
    },
    stopRemoteSession: () => calls.push('stop-remote'),
    logInfo: (message) => calls.push(`info:${message}`),
    logError: (message) => calls.push(`error:${message}`),
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    clearSetupWindow: () => calls.push('clear'),
    setSetupWindow: () => calls.push('set-window'),
    registerSetupIpcHandler: () => {
      calls.push('register-ipc');
      return () => calls.push('unregister-ipc');
    },
    encodeURIComponent: (value) => encodeURIComponent(value),
    defaultServerUrl: 'http://127.0.0.1:8096',
    hasStoredSession: () => true,
  })();

  assert.equal(deps.maybeFocusExistingSetupWindow(), false);
  assert.deepEqual(deps.getResolvedJellyfinConfig(), {
    serverUrl: 'http://127.0.0.1:8096',
    username: 'alice',
  });
  assert.equal(deps.buildSetupFormHtml(expectedState), '<html></html>');
  assert.deepEqual(capturedBuildState, expectedState);
  const setupUrl = 'subminer://jellyfin-setup?x=1';
  assert.deepEqual(deps.parseSubmissionUrl(setupUrl), {
    action: 'login',
    server: 's',
    username: 'u',
    password: 'p',
  });
  assert.equal(capturedParseUrl, setupUrl);
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
  deps.persistAuthenticatedSession?.(
    {
      serverUrl: 'http://127.0.0.1:8096',
      username: 'alice',
      accessToken: 'token',
      userId: 'uid',
    },
    deps.getJellyfinClientInfo(),
  );
  await deps.restartRemoteSession?.();
  deps.stopRemoteSession?.();
  deps.logInfo('ok');
  deps.logError('bad', null);
  deps.showMpvOsd('toast');
  deps.clearSetupWindow();
  deps.setSetupWindow({} as never);
  const unregister = deps.registerSetupIpcHandler?.(async () => ({ handled: true }));
  unregister?.();
  assert.equal(deps.encodeURIComponent('a b'), 'a%20b');
  assert.equal(deps.defaultServerUrl, 'http://127.0.0.1:8096');
  assert.equal(deps.hasStoredSession(), true);
  assert.deepEqual(calls, [
    'save',
    'clear-session',
    'patch',
    'persist',
    'restart-remote',
    'stop-remote',
    'info:ok',
    'error:bad',
    'osd:toast',
    'clear',
    'set-window',
    'register-ipc',
    'unregister-ipc',
  ]);
});

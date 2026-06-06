import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildJellyfinSetupSubmissionUrl,
  buildJellyfinSetupFormHtml,
  buildJellyfinSetupViewState,
  createHandleJellyfinSetupWindowClosedHandler,
  createHandleJellyfinSetupNavigationHandler,
  createHandleJellyfinSetupSubmissionHandler,
  createHandleJellyfinSetupWindowOpenedHandler,
  createMaybeFocusExistingJellyfinSetupWindowHandler,
  createOpenJellyfinSetupWindowHandler,
  normalizeJellyfinSetupIpcSubmission,
  parseJellyfinSetupSubmissionUrl,
} from './jellyfin-setup-window';

test('buildJellyfinSetupFormHtml escapes default values', () => {
  const html = buildJellyfinSetupFormHtml({
    selectedServerUrl: 'http://host/"x"',
    username: 'user"name',
    hasStoredSession: true,
    statusMessage: 'Ready "now"',
    statusKind: 'success',
  });
  assert.ok(html.includes('http://host/&quot;x&quot;'));
  assert.ok(html.includes('user&quot;name'));
  assert.ok(html.includes('Ready &quot;now&quot;'));
  assert.ok(html.includes('Logout'));
  assert.equal(html.includes('Server presets'), false);
  assert.equal(html.includes('serverSelect'), false);
  assert.ok(html.includes('window.subminerJellyfinSetup'));
  assert.ok(html.includes('Logging in to Jellyfin'));
  assert.ok(html.includes('subminer://jellyfin-setup?'));
  assert.equal(html.includes('params.set("password", passwordValue)'), false);
  assert.ok(html.includes('window.__subminerJellyfinPassword = passwordValue'));
});

test('buildJellyfinSetupViewState prefills configured server URL', () => {
  const state = buildJellyfinSetupViewState({
    config: {
      serverUrl: ' http://configured:8096/ ',
      username: 'alice',
      recentServers: ['http://recent:8096', 'http://configured:8096', ''],
    },
    defaultServerUrl: 'http://127.0.0.1:8096',
    hasStoredSession: false,
  });

  assert.equal(state.selectedServerUrl, 'http://configured:8096');
  assert.equal(state.username, 'alice');
  assert.equal(state.statusKind, 'idle');
});

test('buildJellyfinSetupViewState falls back to recent server URL', () => {
  const state = buildJellyfinSetupViewState({
    config: {
      serverUrl: '',
      username: 'alice',
      recentServers: ['http://recent:8096'],
    },
    defaultServerUrl: 'http://127.0.0.1:8096',
    hasStoredSession: false,
  });

  assert.equal(state.selectedServerUrl, 'http://recent:8096');
});

test('maybe focus jellyfin setup window no-ops without window', () => {
  const handler = createMaybeFocusExistingJellyfinSetupWindowHandler({
    getSetupWindow: () => null,
  });
  const handled = handler();
  assert.equal(handled, false);
});

test('parseJellyfinSetupSubmissionUrl parses setup url parameters', () => {
  const parsed = parseJellyfinSetupSubmissionUrl(
    'subminer://jellyfin-setup?action=login&server=http%3A%2F%2Flocalhost&username=a&password=b',
  );
  assert.deepEqual(parsed, {
    action: 'login',
    server: 'http://localhost',
    username: 'a',
    password: 'b',
  });
  assert.deepEqual(parseJellyfinSetupSubmissionUrl('subminer://jellyfin-setup?action=logout'), {
    action: 'logout',
    server: '',
    username: '',
    password: '',
  });
  assert.deepEqual(parseJellyfinSetupSubmissionUrl('subminer://jellyfin-setup?action=done'), {
    action: 'done',
    server: '',
    username: '',
    password: '',
  });
  assert.equal(parseJellyfinSetupSubmissionUrl('https://example.com'), null);
});

test('jellyfin setup ipc submissions normalize to password-free setup urls', () => {
  const submission = normalizeJellyfinSetupIpcSubmission({
    action: 'login',
    server: 'http://localhost:8096',
    username: 'alice',
    password: 'secret',
  });

  assert.deepEqual(submission, {
    action: 'login',
    server: 'http://localhost:8096',
    username: 'alice',
    password: 'secret',
  });
  if (!submission) {
    throw new Error('missing normalized submission');
  }
  const setupUrl = buildJellyfinSetupSubmissionUrl(submission);
  assert.equal(
    setupUrl,
    'subminer://jellyfin-setup?action=login&server=http%3A%2F%2Flocalhost%3A8096&username=alice',
  );
  assert.equal(setupUrl.includes('secret'), false);
  assert.deepEqual(normalizeJellyfinSetupIpcSubmission({ action: 'done' }), {
    action: 'done',
    server: '',
    username: '',
    password: '',
  });
  assert.equal(normalizeJellyfinSetupIpcSubmission('bad'), null);
});

test('createHandleJellyfinSetupSubmissionHandler applies successful login', async () => {
  const calls: string[] = [];
  let patchPayload: unknown = null;
  let savedSession: unknown = null;
  let authPassword = '';
  const handler = createHandleJellyfinSetupSubmissionHandler({
    parseSubmissionUrl: (rawUrl) => parseJellyfinSetupSubmissionUrl(rawUrl),
    authenticateWithPassword: async (_server, _username, password) => {
      authPassword = password;
      return {
        serverUrl: 'http://localhost',
        username: 'user',
        accessToken: 'token',
        userId: 'uid',
      };
    },
    getJellyfinClientInfo: () => ({
      clientName: 'SubMiner',
      clientVersion: '1.0',
      deviceId: 'did',
    }),
    saveStoredSession: (session) => {
      savedSession = session;
      calls.push('save');
    },
    clearStoredSession: () => calls.push('clear'),
    patchJellyfinConfig: (session) => {
      patchPayload = session;
      calls.push('patch');
    },
    restartRemoteSession: async () => {
      calls.push('restart-remote');
    },
    logInfo: () => calls.push('info'),
    logError: () => calls.push('error'),
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    closeSetupWindow: () => calls.push('close'),
    reloadSetupWindow: () => calls.push('reload'),
  });

  const handled = await handler(
    'subminer://jellyfin-setup?action=login&server=http%3A%2F%2Flocalhost&username=a',
    'b',
  );
  assert.equal(handled, true);
  assert.deepEqual(calls, [
    'save',
    'patch',
    'restart-remote',
    'info',
    'osd:Jellyfin login success',
    'reload',
  ]);
  assert.equal(authPassword, 'b');
  assert.deepEqual(savedSession, { accessToken: 'token', userId: 'uid' });
  assert.deepEqual(patchPayload, {
    serverUrl: 'http://localhost',
    username: 'user',
    accessToken: 'token',
    userId: 'uid',
  });
});

test('createHandleJellyfinSetupSubmissionHandler reports failure to OSD', async () => {
  const calls: string[] = [];
  const handler = createHandleJellyfinSetupSubmissionHandler({
    parseSubmissionUrl: (rawUrl) => parseJellyfinSetupSubmissionUrl(rawUrl),
    authenticateWithPassword: async () => {
      throw new Error('bad credentials');
    },
    getJellyfinClientInfo: () => ({
      clientName: 'SubMiner',
      clientVersion: '1.0',
      deviceId: 'did',
    }),
    saveStoredSession: () => calls.push('save'),
    clearStoredSession: () => calls.push('clear'),
    patchJellyfinConfig: () => calls.push('patch'),
    logInfo: () => calls.push('info'),
    logError: () => calls.push('error'),
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    closeSetupWindow: () => calls.push('close'),
    reloadSetupWindow: (_state) => calls.push('reload'),
  });

  const handled = await handler(
    'subminer://jellyfin-setup?action=login&server=http%3A%2F%2Flocalhost&username=a&password=b',
  );
  assert.equal(handled, true);
  assert.deepEqual(calls, ['error', 'osd:Jellyfin login failed: bad credentials', 'reload']);
});

test('createHandleJellyfinSetupSubmissionHandler reports logout failure inline', async () => {
  const calls: string[] = [];
  let reloadState: unknown = null;
  const handler = createHandleJellyfinSetupSubmissionHandler({
    parseSubmissionUrl: (rawUrl) => parseJellyfinSetupSubmissionUrl(rawUrl),
    authenticateWithPassword: async () => {
      throw new Error('should not authenticate');
    },
    getJellyfinClientInfo: () => ({
      clientName: 'SubMiner',
      clientVersion: '1.0',
      deviceId: 'did',
    }),
    saveStoredSession: () => calls.push('save'),
    clearStoredSession: () => {
      throw new Error('logout failed');
    },
    patchJellyfinConfig: () => calls.push('patch'),
    logInfo: () => calls.push('info'),
    logError: (message) => calls.push(`error:${message}`),
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    closeSetupWindow: () => calls.push('close'),
    reloadSetupWindow: (state) => {
      reloadState = state;
      calls.push('reload');
    },
  });

  assert.equal(await handler('subminer://jellyfin-setup?action=logout'), true);
  assert.deepEqual(calls, [
    'error:Jellyfin logout failed',
    'osd:Jellyfin logout failed: logout failed',
    'reload',
  ]);
  assert.deepEqual(reloadState, {
    statusMessage: 'logout failed',
    statusKind: 'error',
  });
});

test('createHandleJellyfinSetupSubmissionHandler ignores concurrent login submissions', async () => {
  const calls: string[] = [];
  type TestSession = {
    serverUrl: string;
    username: string;
    accessToken: string;
    userId: string;
  };
  let finishAuth: ((session: TestSession) => void) | undefined;
  const handler = createHandleJellyfinSetupSubmissionHandler({
    parseSubmissionUrl: (rawUrl) => parseJellyfinSetupSubmissionUrl(rawUrl),
    authenticateWithPassword: async () =>
      new Promise<TestSession>((resolve) => {
        finishAuth = resolve;
      }),
    getJellyfinClientInfo: () => ({
      clientName: 'SubMiner',
      clientVersion: '1.0',
      deviceId: 'did',
    }),
    saveStoredSession: () => calls.push('save'),
    clearStoredSession: () => calls.push('clear'),
    patchJellyfinConfig: () => calls.push('patch'),
    logInfo: () => calls.push('info'),
    logError: () => calls.push('error'),
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    closeSetupWindow: () => calls.push('close'),
    reloadSetupWindow: (state) => calls.push(`reload:${state?.statusKind || 'none'}`),
  });

  const first = handler(
    'subminer://jellyfin-setup?action=login&server=http%3A%2F%2Flocalhost&username=a',
    'first',
  );
  const second = await handler(
    'subminer://jellyfin-setup?action=login&server=http%3A%2F%2Flocalhost&username=a',
    'second',
  );

  assert.equal(second, true);
  const resolveAuth = finishAuth;
  if (!resolveAuth) {
    throw new Error('missing auth resolver');
  }
  resolveAuth({
    serverUrl: 'http://localhost',
    username: 'a',
    accessToken: 'token',
    userId: 'uid',
  });
  assert.equal(await first, true);
  assert.deepEqual(calls, [
    'osd:Jellyfin login already in progress',
    'reload:loading',
    'save',
    'patch',
    'info',
    'osd:Jellyfin login success',
    'reload:success',
  ]);
});

test('createHandleJellyfinSetupSubmissionHandler handles logout and done', async () => {
  const calls: string[] = [];
  const handler = createHandleJellyfinSetupSubmissionHandler({
    parseSubmissionUrl: (rawUrl) => parseJellyfinSetupSubmissionUrl(rawUrl),
    authenticateWithPassword: async () => {
      throw new Error('should not authenticate');
    },
    getJellyfinClientInfo: () => ({
      clientName: 'SubMiner',
      clientVersion: '1.0',
      deviceId: 'did',
    }),
    saveStoredSession: () => calls.push('save'),
    clearStoredSession: () => calls.push('clear'),
    patchJellyfinConfig: () => calls.push('patch'),
    stopRemoteSession: () => calls.push('stop-remote'),
    logInfo: (message) => calls.push(message),
    logError: () => calls.push('error'),
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    closeSetupWindow: () => calls.push('close'),
    reloadSetupWindow: () => calls.push('reload'),
  });

  assert.equal(await handler('subminer://jellyfin-setup?action=logout'), true);
  assert.equal(await handler('subminer://jellyfin-setup?action=done'), true);
  assert.deepEqual(calls, [
    'clear',
    'stop-remote',
    'Cleared stored Jellyfin auth session.',
    'osd:Jellyfin logged out',
    'reload',
    'close',
  ]);
});

test('createHandleJellyfinSetupNavigationHandler ignores unrelated urls', () => {
  const handleNavigation = createHandleJellyfinSetupNavigationHandler({
    setupSchemePrefix: 'subminer://jellyfin-setup',
    handleSubmission: async () => {},
    logError: () => {},
  });
  let prevented = false;
  const handled = handleNavigation({
    url: 'https://example.com',
    preventDefault: () => {
      prevented = true;
    },
  });
  assert.equal(handled, false);
  assert.equal(prevented, false);
});

test('createHandleJellyfinSetupNavigationHandler intercepts setup urls', async () => {
  const submittedUrls: string[] = [];
  const handleNavigation = createHandleJellyfinSetupNavigationHandler({
    setupSchemePrefix: 'subminer://jellyfin-setup',
    handleSubmission: async (rawUrl) => {
      submittedUrls.push(rawUrl);
    },
    logError: () => {},
  });
  let prevented = false;
  const handled = handleNavigation({
    url: 'subminer://jellyfin-setup?server=http%3A%2F%2F127.0.0.1%3A8096',
    preventDefault: () => {
      prevented = true;
    },
  });
  await Promise.resolve();
  assert.equal(handled, true);
  assert.equal(prevented, true);
  assert.equal(submittedUrls.length, 1);
});

test('createHandleJellyfinSetupWindowClosedHandler clears setup window ref', () => {
  let cleared = false;
  const handler = createHandleJellyfinSetupWindowClosedHandler({
    clearSetupWindow: () => {
      cleared = true;
    },
  });
  handler();
  assert.equal(cleared, true);
});

test('createHandleJellyfinSetupWindowOpenedHandler sets setup window ref', () => {
  let set = false;
  const handler = createHandleJellyfinSetupWindowOpenedHandler({
    setSetupWindow: () => {
      set = true;
    },
  });
  handler();
  assert.equal(set, true);
});

test('createOpenJellyfinSetupWindowHandler no-ops when existing setup window is focused', () => {
  const calls: string[] = [];
  const handler = createOpenJellyfinSetupWindowHandler({
    maybeFocusExistingSetupWindow: () => {
      calls.push('focus-existing');
      return true;
    },
    createSetupWindow: () => {
      calls.push('create-window');
      throw new Error('should not create');
    },
    getResolvedJellyfinConfig: () => ({}),
    buildSetupFormHtml: () => '<html></html>',
    parseSubmissionUrl: () => null,
    authenticateWithPassword: async () => {
      throw new Error('should not auth');
    },
    getJellyfinClientInfo: () => ({
      clientName: 'SubMiner',
      clientVersion: '1.0',
      deviceId: 'did',
    }),
    saveStoredSession: () => {},
    patchJellyfinConfig: () => {},
    logInfo: () => {},
    logError: () => {},
    showMpvOsd: () => {},
    clearSetupWindow: () => {},
    setSetupWindow: () => {},
    clearStoredSession: () => {},
    encodeURIComponent: (value) => value,
    defaultServerUrl: 'http://127.0.0.1:8096',
    hasStoredSession: () => false,
  });

  handler();
  assert.deepEqual(calls, ['focus-existing']);
});

test('createOpenJellyfinSetupWindowHandler wires navigation, load, and window lifecycle', async () => {
  let willNavigateHandler: ((event: { preventDefault: () => void }, url: string) => void) | null =
    null;
  let closedHandler: (() => void) | null = null;
  let prevented = false;
  const calls: string[] = [];
  const fakeWindow = {
    focus: () => {},
    webContents: {
      on: (
        event: 'will-navigate',
        handler: (event: { preventDefault: () => void }, url: string) => void,
      ) => {
        if (event === 'will-navigate') {
          willNavigateHandler = handler;
        }
      },
      executeJavaScript: async () => 'pass',
    },
    loadURL: (url: string) => {
      calls.push(`load:${url.startsWith('data:text/html;charset=utf-8,') ? 'data-url' : 'other'}`);
    },
    on: (event: 'closed', handler: () => void) => {
      if (event === 'closed') {
        closedHandler = handler;
      }
    },
    isDestroyed: () => false,
    close: () => calls.push('close'),
  };

  const handler = createOpenJellyfinSetupWindowHandler({
    maybeFocusExistingSetupWindow: () => false,
    createSetupWindow: () => fakeWindow,
    getResolvedJellyfinConfig: () => ({
      serverUrl: 'http://localhost:8096',
      username: 'alice',
      recentServers: [],
    }),
    buildSetupFormHtml: (state) => `<html>${state.selectedServerUrl}|${state.username}</html>`,
    parseSubmissionUrl: (rawUrl) => parseJellyfinSetupSubmissionUrl(rawUrl),
    authenticateWithPassword: async (_server, _username, password) => {
      calls.push(`password:${password}`);
      return {
        serverUrl: 'http://localhost:8096',
        username: 'alice',
        accessToken: 'token',
        userId: 'uid',
      };
    },
    getJellyfinClientInfo: () => ({
      clientName: 'SubMiner',
      clientVersion: '1.0',
      deviceId: 'did',
    }),
    saveStoredSession: () => calls.push('save'),
    clearStoredSession: () => calls.push('clear'),
    patchJellyfinConfig: () => calls.push('patch'),
    logInfo: () => calls.push('info'),
    logError: () => calls.push('error'),
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    clearSetupWindow: () => calls.push('clear-window'),
    setSetupWindow: () => calls.push('set-window'),
    encodeURIComponent: (value) => encodeURIComponent(value),
    defaultServerUrl: 'http://127.0.0.1:8096',
    hasStoredSession: () => true,
  });

  handler();
  assert.ok(willNavigateHandler);
  assert.ok(closedHandler);
  assert.deepEqual(calls.slice(0, 2), ['load:data-url', 'set-window']);

  const navHandler = willNavigateHandler as
    | ((event: { preventDefault: () => void }, url: string) => void)
    | null;
  if (!navHandler) {
    throw new Error('missing will-navigate handler');
  }
  navHandler(
    {
      preventDefault: () => {
        prevented = true;
      },
    },
    'subminer://jellyfin-setup?action=login&server=http%3A%2F%2Flocalhost&username=alice',
  );
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.equal(prevented, true);
  assert.ok(calls.includes('password:pass'));
  assert.ok(calls.includes('save'));
  assert.ok(calls.includes('patch'));
  assert.ok(calls.includes('osd:Jellyfin login success'));
  assert.ok(calls.includes('load:data-url'));

  const onClosed = closedHandler as (() => void) | null;
  if (!onClosed) {
    throw new Error('missing closed handler');
  }
  onClosed();
  assert.ok(calls.includes('clear-window'));
});

test('createOpenJellyfinSetupWindowHandler handles ipc bridge submissions', async () => {
  const bridge: { handler?: (payload: unknown) => Promise<{ handled: boolean }> } = {};
  let closedHandler: (() => void) | null = null;
  const calls: string[] = [];
  const fakeWindow = {
    focus: () => {},
    webContents: {
      on: () => {},
      executeJavaScript: async () => {
        throw new Error('bridge path should not read from page');
      },
    },
    loadURL: () => {
      calls.push('load');
    },
    on: (event: 'closed', handler: () => void) => {
      if (event === 'closed') {
        closedHandler = handler;
      }
    },
    isDestroyed: () => false,
    close: () => calls.push('close'),
  };

  const handler = createOpenJellyfinSetupWindowHandler({
    maybeFocusExistingSetupWindow: () => false,
    createSetupWindow: () => fakeWindow,
    getResolvedJellyfinConfig: () => ({
      serverUrl: 'http://localhost:8096',
      username: 'alice',
      recentServers: [],
    }),
    buildSetupFormHtml: () => '<html></html>',
    parseSubmissionUrl: (rawUrl) => parseJellyfinSetupSubmissionUrl(rawUrl),
    authenticateWithPassword: async (_server, _username, password) => {
      calls.push(`password:${password}`);
      return {
        serverUrl: 'http://localhost:8096',
        username: 'alice',
        accessToken: 'token',
        userId: 'uid',
      };
    },
    getJellyfinClientInfo: () => ({
      clientName: 'SubMiner',
      clientVersion: '1.0',
      deviceId: 'did',
    }),
    saveStoredSession: () => calls.push('save'),
    clearStoredSession: () => calls.push('clear'),
    patchJellyfinConfig: () => calls.push('patch'),
    logInfo: () => calls.push('info'),
    logError: () => calls.push('error'),
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    clearSetupWindow: () => calls.push('clear-window'),
    setSetupWindow: () => calls.push('set-window'),
    registerSetupIpcHandler: (nextHandler) => {
      bridge.handler = nextHandler;
      calls.push('register-ipc');
      return () => calls.push('unregister-ipc');
    },
    encodeURIComponent: (value) => encodeURIComponent(value),
    defaultServerUrl: 'http://127.0.0.1:8096',
    hasStoredSession: () => false,
  });

  handler();
  const bridgeHandler = bridge.handler;
  if (!bridgeHandler) {
    throw new Error('missing bridge handler');
  }
  assert.deepEqual(await bridgeHandler('bad'), {
    handled: false,
    statusMessage: 'Invalid Jellyfin setup request.',
    statusKind: 'error',
  });
  assert.equal(calls.includes('password:'), false);

  calls.length = 0;
  assert.deepEqual(
    await bridgeHandler({
      action: 'login',
      server: 'http://localhost:8096',
      username: 'alice',
      password: 'secret',
    }),
    { handled: true },
  );
  assert.ok(calls.includes('password:secret'));
  assert.ok(calls.includes('save'));
  assert.ok(calls.includes('patch'));

  const onClosed = closedHandler as (() => void) | null;
  if (!onClosed) {
    throw new Error('missing closed handler');
  }
  onClosed();
  assert.ok(calls.includes('unregister-ipc'));
  assert.ok(calls.includes('clear-window'));
});

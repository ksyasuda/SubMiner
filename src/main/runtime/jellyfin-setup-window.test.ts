import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildJellyfinSetupFormHtml,
  buildJellyfinSetupViewState,
  createHandleJellyfinSetupWindowClosedHandler,
  createHandleJellyfinSetupNavigationHandler,
  createHandleJellyfinSetupSubmissionHandler,
  createHandleJellyfinSetupWindowOpenedHandler,
  createMaybeFocusExistingJellyfinSetupWindowHandler,
  createOpenJellyfinSetupWindowHandler,
  parseJellyfinSetupSubmissionUrl,
} from './jellyfin-setup-window';

test('buildJellyfinSetupFormHtml escapes default values', () => {
  const html = buildJellyfinSetupFormHtml({
    servers: [
      {
        serverUrl: 'http://host/"x"',
        label: 'Configured "Server"',
        source: 'config',
      },
    ],
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
  assert.ok(html.includes('subminer://jellyfin-setup?'));
});

test('buildJellyfinSetupViewState composes config, recent, and default servers', () => {
  const state = buildJellyfinSetupViewState({
    config: {
      serverUrl: ' http://configured:8096/ ',
      username: 'alice',
      recentServers: ['http://recent:8096', 'http://configured:8096', ''],
    },
    defaultServerUrl: 'http://127.0.0.1:8096',
    hasStoredSession: false,
  });

  assert.deepEqual(
    state.servers.map((server) => [server.serverUrl, server.source]),
    [
      ['http://configured:8096', 'config'],
      ['http://recent:8096', 'recent'],
      ['http://127.0.0.1:8096', 'default'],
    ],
  );
  assert.equal(state.selectedServerUrl, 'http://configured:8096');
  assert.equal(state.username, 'alice');
  assert.equal(state.statusKind, 'idle');
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

test('createHandleJellyfinSetupSubmissionHandler applies successful login', async () => {
  const calls: string[] = [];
  let patchPayload: unknown = null;
  let savedSession: unknown = null;
  const handler = createHandleJellyfinSetupSubmissionHandler({
    parseSubmissionUrl: (rawUrl) => parseJellyfinSetupSubmissionUrl(rawUrl),
    authenticateWithPassword: async () => ({
      serverUrl: 'http://localhost',
      username: 'user',
      accessToken: 'token',
      userId: 'uid',
    }),
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
    logInfo: () => calls.push('info'),
    logError: () => calls.push('error'),
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    closeSetupWindow: () => calls.push('close'),
    reloadSetupWindow: () => calls.push('reload'),
  });

  const handled = await handler(
    'subminer://jellyfin-setup?action=login&server=http%3A%2F%2Flocalhost&username=a&password=b',
  );
  assert.equal(handled, true);
  assert.deepEqual(calls, ['save', 'patch', 'info', 'osd:Jellyfin login success', 'reload']);
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
    authenticateWithPassword: async () => ({
      serverUrl: 'http://localhost:8096',
      username: 'alice',
      accessToken: 'token',
      userId: 'uid',
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
    'subminer://jellyfin-setup?action=login&server=http%3A%2F%2Flocalhost&username=alice&password=pass',
  );
  await Promise.resolve();

  assert.equal(prevented, true);
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

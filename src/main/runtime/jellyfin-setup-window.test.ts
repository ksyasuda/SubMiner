import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildJellyfinSetupFormHtml,
  createHandleJellyfinSetupWindowClosedHandler,
  createHandleJellyfinSetupNavigationHandler,
  createHandleJellyfinSetupSubmissionHandler,
  createHandleJellyfinSetupWindowOpenedHandler,
  createMaybeFocusExistingJellyfinSetupWindowHandler,
  createOpenJellyfinSetupWindowHandler,
  parseJellyfinSetupSubmissionUrl,
} from './jellyfin-setup-window';

test('buildJellyfinSetupFormHtml escapes default values', () => {
  const html = buildJellyfinSetupFormHtml('http://host/"x"', 'user"name');
  assert.ok(html.includes('http://host/&quot;x&quot;'));
  assert.ok(html.includes('user&quot;name'));
  assert.ok(html.includes('subminer://jellyfin-setup?'));
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
    'subminer://jellyfin-setup?server=http%3A%2F%2Flocalhost&username=a&password=b',
  );
  assert.deepEqual(parsed, {
    server: 'http://localhost',
    username: 'a',
    password: 'b',
  });
  assert.equal(parseJellyfinSetupSubmissionUrl('https://example.com'), null);
});

test('createHandleJellyfinSetupSubmissionHandler applies successful login', async () => {
  const calls: string[] = [];
  const handler = createHandleJellyfinSetupSubmissionHandler({
    parseSubmissionUrl: (rawUrl) => parseJellyfinSetupSubmissionUrl(rawUrl),
    authenticateWithPassword: async () => ({
      serverUrl: 'http://localhost',
      username: 'user',
      accessToken: 'token',
      userId: 'uid',
    }),
    getJellyfinClientInfo: () => ({ clientName: 'SubMiner', clientVersion: '1.0', deviceId: 'did' }),
    patchJellyfinConfig: () => calls.push('patch'),
    logInfo: () => calls.push('info'),
    logError: () => calls.push('error'),
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    closeSetupWindow: () => calls.push('close'),
  });

  const handled = await handler(
    'subminer://jellyfin-setup?server=http%3A%2F%2Flocalhost&username=a&password=b',
  );
  assert.equal(handled, true);
  assert.deepEqual(calls, ['patch', 'info', 'osd:Jellyfin login success', 'close']);
});

test('createHandleJellyfinSetupSubmissionHandler reports failure to OSD', async () => {
  const calls: string[] = [];
  const handler = createHandleJellyfinSetupSubmissionHandler({
    parseSubmissionUrl: (rawUrl) => parseJellyfinSetupSubmissionUrl(rawUrl),
    authenticateWithPassword: async () => {
      throw new Error('bad credentials');
    },
    getJellyfinClientInfo: () => ({ clientName: 'SubMiner', clientVersion: '1.0', deviceId: 'did' }),
    patchJellyfinConfig: () => calls.push('patch'),
    logInfo: () => calls.push('info'),
    logError: () => calls.push('error'),
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    closeSetupWindow: () => calls.push('close'),
  });

  const handled = await handler(
    'subminer://jellyfin-setup?server=http%3A%2F%2Flocalhost&username=a&password=b',
  );
  assert.equal(handled, true);
  assert.deepEqual(calls, ['error', 'osd:Jellyfin login failed: bad credentials']);
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
    getJellyfinClientInfo: () => ({ clientName: 'SubMiner', clientVersion: '1.0', deviceId: 'did' }),
    patchJellyfinConfig: () => {},
    logInfo: () => {},
    logError: () => {},
    showMpvOsd: () => {},
    clearSetupWindow: () => {},
    setSetupWindow: () => {},
    encodeURIComponent: (value) => value,
  });

  handler();
  assert.deepEqual(calls, ['focus-existing']);
});

test('createOpenJellyfinSetupWindowHandler wires navigation, load, and window lifecycle', async () => {
  let willNavigateHandler: ((event: { preventDefault: () => void }, url: string) => void) | null = null;
  let closedHandler: (() => void) | null = null;
  let prevented = false;
  const calls: string[] = [];
  const fakeWindow = {
    focus: () => {},
    webContents: {
      on: (event: 'will-navigate', handler: (event: { preventDefault: () => void }, url: string) => void) => {
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
    getResolvedJellyfinConfig: () => ({ serverUrl: 'http://localhost:8096', username: 'alice' }),
    buildSetupFormHtml: (server, username) => `<html>${server}|${username}</html>`,
    parseSubmissionUrl: (rawUrl) => parseJellyfinSetupSubmissionUrl(rawUrl),
    authenticateWithPassword: async () => ({
      serverUrl: 'http://localhost:8096',
      username: 'alice',
      accessToken: 'token',
      userId: 'uid',
    }),
    getJellyfinClientInfo: () => ({ clientName: 'SubMiner', clientVersion: '1.0', deviceId: 'did' }),
    patchJellyfinConfig: () => calls.push('patch'),
    logInfo: () => calls.push('info'),
    logError: () => calls.push('error'),
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    clearSetupWindow: () => calls.push('clear-window'),
    setSetupWindow: () => calls.push('set-window'),
    encodeURIComponent: (value) => encodeURIComponent(value),
  });

  handler();
  assert.ok(willNavigateHandler);
  assert.ok(closedHandler);
  assert.deepEqual(calls.slice(0, 2), ['load:data-url', 'set-window']);

  const navHandler = willNavigateHandler as ((event: { preventDefault: () => void }, url: string) => void) | null;
  if (!navHandler) {
    throw new Error('missing will-navigate handler');
  }
  navHandler(
    {
      preventDefault: () => {
        prevented = true;
      },
    },
    'subminer://jellyfin-setup?server=http%3A%2F%2Flocalhost&username=alice&password=pass',
  );
  await Promise.resolve();

  assert.equal(prevented, true);
  assert.ok(calls.includes('patch'));
  assert.ok(calls.includes('osd:Jellyfin login success'));
  assert.ok(calls.includes('close'));

  const onClosed = closedHandler as (() => void) | null;
  if (!onClosed) {
    throw new Error('missing closed handler');
  }
  onClosed();
  assert.ok(calls.includes('clear-window'));
});

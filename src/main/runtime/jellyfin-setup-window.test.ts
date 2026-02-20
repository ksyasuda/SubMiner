import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildJellyfinSetupFormHtml,
  createHandleJellyfinSetupWindowClosedHandler,
  createHandleJellyfinSetupNavigationHandler,
  createHandleJellyfinSetupSubmissionHandler,
  createHandleJellyfinSetupWindowOpenedHandler,
  createMaybeFocusExistingJellyfinSetupWindowHandler,
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

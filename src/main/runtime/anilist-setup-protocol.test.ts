import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createConsumeAnilistSetupTokenFromUrlHandler,
  createHandleAnilistSetupProtocolUrlHandler,
  createNotifyAnilistSetupHandler,
  createRegisterSubminerProtocolClientHandler,
} from './anilist-setup-protocol';

test('createNotifyAnilistSetupHandler routes AniList setup messages through configured notifications', () => {
  const calls: string[] = [];
  const notify = createNotifyAnilistSetupHandler({
    showConfiguredNotification: (title, payload) =>
      calls.push(`configured:${title}:${payload.kind}:${payload.message}`),
    logInfo: () => calls.push('log'),
  });
  notify('AniList login success');
  assert.deepEqual(calls, [
    'configured:SubMiner AniList:success:AniList login success',
    'log',
  ]);
});

test('createConsumeAnilistSetupTokenFromUrlHandler delegates with deps', () => {
  const consume = createConsumeAnilistSetupTokenFromUrlHandler({
    consumeAnilistSetupCallbackUrl: (input) => input.rawUrl.includes('access_token=ok'),
    saveToken: () => {},
    setCachedToken: () => {},
    setResolvedState: () => {},
    setSetupPageOpened: () => {},
    onSuccess: () => {},
    closeWindow: () => {},
  });
  assert.equal(consume('subminer://anilist-setup?access_token=ok'), true);
  assert.equal(consume('subminer://anilist-setup'), false);
});

test('createHandleAnilistSetupProtocolUrlHandler validates scheme and logs missing token', () => {
  const warnings: string[] = [];
  const handleProtocolUrl = createHandleAnilistSetupProtocolUrlHandler({
    consumeAnilistSetupTokenFromUrl: () => false,
    logWarn: (message) => warnings.push(message),
  });

  assert.equal(handleProtocolUrl('https://example.com'), false);
  assert.equal(handleProtocolUrl('subminer://anilist-setup'), true);
  assert.deepEqual(warnings, ['AniList setup protocol URL missing access token']);
});

test('createRegisterSubminerProtocolClientHandler registers default app entry', () => {
  const calls: string[] = [];
  const register = createRegisterSubminerProtocolClientHandler({
    isDefaultApp: () => true,
    getArgv: () => ['electron', './entry.js'],
    execPath: '/usr/local/bin/electron',
    resolvePath: (value) => `/resolved/${value}`,
    setAsDefaultProtocolClient: (_scheme, _path, args) => {
      calls.push(`register:${String(args?.[0])}`);
      return true;
    },
    logDebug: (message) => calls.push(`debug:${message}`),
  });

  register();
  assert.deepEqual(calls, ['register:/resolved/./entry.js']);
});

test('createRegisterSubminerProtocolClientHandler keeps unsupported registration at debug level', () => {
  const calls: string[] = [];
  const register = createRegisterSubminerProtocolClientHandler({
    isDefaultApp: () => false,
    getArgv: () => ['SubMiner.AppImage'],
    execPath: '/tmp/SubMiner.AppImage',
    resolvePath: (value) => value,
    setAsDefaultProtocolClient: () => false,
    logDebug: (message) => calls.push(`debug:${message}`),
  });

  register();
  assert.deepEqual(calls, [
    'debug:Failed to register default protocol handler for subminer:// URLs',
  ]);
});

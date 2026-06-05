import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createConsumeAnilistSetupTokenFromUrlHandler,
  createHandleAnilistSetupProtocolUrlHandler,
  createNotifyAnilistSetupHandler,
  createRegisterSubminerProtocolClientHandler,
} from './anilist-setup-protocol';

test('createNotifyAnilistSetupHandler sends OSD when mpv client exists', () => {
  const calls: string[] = [];
  const notify = createNotifyAnilistSetupHandler({
    hasMpvClient: () => true,
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    showDesktopNotification: () => calls.push('desktop'),
    logInfo: () => calls.push('log'),
  });
  notify('AniList login success');
  assert.deepEqual(calls, ['osd:AniList login success']);
});

test('createNotifyAnilistSetupHandler routes through configured notification surfaces', () => {
  const calls: string[] = [];
  const notify = createNotifyAnilistSetupHandler({
    getNotificationType: () => 'both',
    hasMpvClient: () => true,
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    showOverlayNotification: (payload) =>
      calls.push(`overlay:${payload.title}:${payload.body}:${payload.variant}`),
    showDesktopNotification: (title, options) => calls.push(`notify:${title}:${options.body}`),
    logInfo: () => calls.push('log'),
  });
  notify('AniList login success');
  assert.deepEqual(calls, [
    'overlay:SubMiner AniList:AniList login success:success',
    'notify:SubMiner AniList:AniList login success',
  ]);
});

test('createConsumeAnilistSetupTokenFromUrlHandler delegates with deps', () => {
  const consume = createConsumeAnilistSetupTokenFromUrlHandler({
    consumeAnilistSetupCallbackUrl: (input) => input.rawUrl.includes('access_token=ok'),
    saveToken: () => true,
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

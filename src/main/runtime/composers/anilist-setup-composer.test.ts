import test from 'node:test';
import assert from 'node:assert/strict';
import { composeAnilistSetupHandlers } from './anilist-setup-composer';

test('composeAnilistSetupHandlers returns callable setup handlers', () => {
  const calls: string[] = [];
  const composed = composeAnilistSetupHandlers({
    notifyDeps: {
      hasMpvClient: () => false,
      showMpvOsd: () => {},
      showDesktopNotification: (title, opts) => {
        calls.push(`notify:${opts.body}`);
      },
      logInfo: () => {},
    },
    consumeTokenDeps: {
      consumeAnilistSetupCallbackUrl: () => false,
      saveToken: () => true,
      setCachedToken: () => {},
      setResolvedState: () => {},
      setSetupPageOpened: () => {},
      onSuccess: () => {},
      closeWindow: () => {},
    },
    handleProtocolDeps: {
      consumeAnilistSetupTokenFromUrl: () => false,
      logWarn: () => {},
    },
    registerProtocolClientDeps: {
      isDefaultApp: () => false,
      getArgv: () => [],
      execPath: process.execPath,
      resolvePath: (value) => value,
      setAsDefaultProtocolClient: () => true,
      logDebug: () => {},
    },
  });

  assert.equal(typeof composed.notifyAnilistSetup, 'function');
  assert.equal(typeof composed.consumeAnilistSetupTokenFromUrl, 'function');
  assert.equal(typeof composed.handleAnilistSetupProtocolUrl, 'function');
  assert.equal(typeof composed.registerSubminerProtocolClient, 'function');

  // notifyAnilistSetup forwards to showDesktopNotification when no MPV client
  composed.notifyAnilistSetup('Setup complete');
  assert.deepEqual(calls, ['notify:Setup complete']);

  // handleAnilistSetupProtocolUrl returns false for non-subminer URLs
  const handled = composed.handleAnilistSetupProtocolUrl('https://other.example.com/');
  assert.equal(handled, false);

  // handleAnilistSetupProtocolUrl returns true for subminer:// URLs
  const handledProtocol = composed.handleAnilistSetupProtocolUrl(
    'subminer://anilist-setup?code=abc',
  );
  assert.equal(handledProtocol, true);
});

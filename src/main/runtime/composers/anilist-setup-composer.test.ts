import test from 'node:test';
import assert from 'node:assert/strict';
import { composeAnilistSetupHandlers } from './anilist-setup-composer';

test('composeAnilistSetupHandlers returns callable setup handlers', () => {
  const composed = composeAnilistSetupHandlers({
    notifyDeps: {
      showConfiguredNotification: () => {},
      logInfo: () => {},
    },
    consumeTokenDeps: {
      consumeAnilistSetupCallbackUrl: () => false,
      saveToken: () => {},
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
});

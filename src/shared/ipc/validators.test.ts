import assert from 'node:assert/strict';
import test from 'node:test';

import { DEFAULT_CONFIG, DEFAULT_KEYBINDINGS } from '../../config/definitions';
import { compileSessionBindings } from '../../core/services/session-bindings';
import { resolveConfiguredShortcuts } from '../../core/utils/shortcut-config';
import { parseSessionActionDispatchRequest } from './validators';

// Regression guard: SESSION_ACTION_IDS in validators.ts is a hand-maintained mirror of the
// SessionActionId union. If a new shortcut-backed action is added to the union/defaults but not to
// the validator allow-list, the renderer's dispatchSessionAction IPC is rejected at runtime (which
// surfaces as a "Renderer error recovered" toast). Compile every default binding and assert the
// validator accepts each one so the two lists can't silently drift apart.
test('every default session-action binding is accepted by parseSessionActionDispatchRequest', () => {
  const { bindings } = compileSessionBindings({
    shortcuts: resolveConfiguredShortcuts(DEFAULT_CONFIG, DEFAULT_CONFIG),
    keybindings: DEFAULT_KEYBINDINGS,
    statsToggleKey: DEFAULT_CONFIG.stats.toggleKey,
    statsMarkWatchedKey: DEFAULT_CONFIG.stats.markWatchedKey,
    platform: 'linux',
    rawConfig: DEFAULT_CONFIG,
  });

  const sessionActions = bindings.filter((binding) => binding.actionType === 'session-action');
  assert.ok(sessionActions.length > 0, 'expected default session-action bindings to exist');

  for (const binding of sessionActions) {
    if (binding.actionType !== 'session-action') continue;
    const request =
      binding.payload === undefined
        ? { actionId: binding.actionId }
        : { actionId: binding.actionId, payload: binding.payload };
    assert.ok(
      parseSessionActionDispatchRequest(request) !== null,
      `validator rejected session action: ${binding.actionId}`,
    );
  }
});

test('toggleNotificationHistory dispatch request is accepted', () => {
  assert.deepEqual(parseSessionActionDispatchRequest({ actionId: 'toggleNotificationHistory' }), {
    actionId: 'toggleNotificationHistory',
  });
});

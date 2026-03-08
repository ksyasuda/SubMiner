import test from 'node:test';
import assert from 'node:assert/strict';
import { DEFAULT_CONFIG, deepCloneConfig } from '../../config';
import {
  buildRestartRequiredConfigMessage,
  createConfigHotReloadAppliedHandler,
  createConfigHotReloadMessageHandler,
} from './config-hot-reload-handlers';

test('createConfigHotReloadAppliedHandler runs all hot-reload effects', () => {
  const config = deepCloneConfig(DEFAULT_CONFIG);
  const calls: string[] = [];
  const ankiPatches: Array<{ enabled: boolean }> = [];

  const applyHotReload = createConfigHotReloadAppliedHandler({
    setKeybindings: () => calls.push('set:keybindings'),
    refreshGlobalAndOverlayShortcuts: () => calls.push('refresh:shortcuts'),
    setSecondarySubMode: (mode) => calls.push(`set:secondary:${mode}`),
    broadcastToOverlayWindows: (channel, payload) =>
      calls.push(`broadcast:${channel}:${typeof payload === 'string' ? payload : 'object'}`),
    applyAnkiRuntimeConfigPatch: (patch) => {
      ankiPatches.push({ enabled: patch.ai });
    },
  });

  applyHotReload(
    {
      hotReloadFields: ['shortcuts', 'secondarySub.defaultMode', 'ankiConnect.ai'],
      restartRequiredFields: [],
    },
    config,
  );

  assert.ok(calls.includes('set:keybindings'));
  assert.ok(calls.includes('refresh:shortcuts'));
  assert.ok(calls.includes(`set:secondary:${config.secondarySub.defaultMode}`));
  assert.ok(calls.some((entry) => entry.startsWith('broadcast:secondary-subtitle:mode:')));
  assert.ok(calls.includes('broadcast:config:hot-reload:object'));
  assert.deepEqual(ankiPatches, [{ enabled: config.ankiConnect.ai.enabled }]);
});

test('createConfigHotReloadAppliedHandler skips optional effects when no hot fields', () => {
  const config = deepCloneConfig(DEFAULT_CONFIG);
  const calls: string[] = [];

  const applyHotReload = createConfigHotReloadAppliedHandler({
    setKeybindings: () => calls.push('set:keybindings'),
    refreshGlobalAndOverlayShortcuts: () => calls.push('refresh:shortcuts'),
    setSecondarySubMode: () => calls.push('set:secondary'),
    broadcastToOverlayWindows: (channel) => calls.push(`broadcast:${channel}`),
    applyAnkiRuntimeConfigPatch: () => calls.push('anki:patch'),
  });

  applyHotReload(
    {
      hotReloadFields: [],
      restartRequiredFields: [],
    },
    config,
  );

  assert.deepEqual(calls, ['set:keybindings']);
});

test('createConfigHotReloadMessageHandler mirrors message to OSD and desktop notification', () => {
  const calls: string[] = [];
  const handleMessage = createConfigHotReloadMessageHandler({
    showMpvOsd: (message) => calls.push(`osd:${message}`),
    showDesktopNotification: (title, options) => calls.push(`notify:${title}:${options.body}`),
  });

  handleMessage('Config reload failed');
  assert.deepEqual(calls, ['osd:Config reload failed', 'notify:SubMiner:Config reload failed']);
});

test('buildRestartRequiredConfigMessage formats changed fields', () => {
  assert.equal(
    buildRestartRequiredConfigMessage(['websocket', 'subtitleStyle']),
    'Config updated; restart required for: websocket, subtitleStyle',
  );
});

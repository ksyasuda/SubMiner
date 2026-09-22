import assert from 'node:assert/strict';
import test from 'node:test';
import { resolveConfig } from '../resolve';
import { DEFAULT_CONFIG } from '../definitions';
import { buildConfigSettingsRegistry } from '../settings/registry';
import { resolveConfiguredShortcuts } from '../../core/utils/shortcut-config';
import {
  compileSessionBindings,
  buildPluginSessionBindingsArtifact,
} from '../../core/services/session-bindings';

function bindings(config: ReturnType<typeof resolveConfig>['resolved']) {
  return compileSessionBindings({
    shortcuts: resolveConfiguredShortcuts(config, DEFAULT_CONFIG),
    keybindings: config.keybindings,
    platform: 'linux',
  });
}

test('subtitle selection is opt-in and enabling it compiles g-s for mpv and the overlay', () => {
  const defaults = resolveConfig({}).resolved;
  assert.equal(defaults.subtitleSelection.enabled, false);
  assert.equal(defaults.shortcuts.openSubtitleSelection, 'g-s');
  const find = (config: typeof defaults) =>
    bindings(config).bindings.find(
      (binding) =>
        binding.actionType === 'session-action' && binding.actionId === 'openSubtitleSelection',
    );
  assert.equal(find(defaults), undefined);
  const enabled = resolveConfig({ subtitleSelection: { enabled: true } }).resolved;
  const binding = find(enabled);
  assert.ok(binding);
  assert.deepEqual(binding.key, { code: 'KeyG-KeyS', modifiers: [] });
  assert.equal(bindings(enabled).warnings.length, 0);
  const artifact = buildPluginSessionBindingsArtifact({
    bindings: [binding],
    warnings: [],
    numericSelectionTimeoutMs: 1000,
  });
  assert.deepEqual(artifact.bindings[0], {
    ...binding,
    cliArgs: ['--session-action', '{"actionId":"openSubtitleSelection"}'],
  });
  enabled.subtitleSelection.enabled = false;
  assert.equal(find(enabled), undefined);
});

test('subtitle selection settings are validated, hot reloadable, and the shortcut can be cleared', () => {
  // @ts-expect-error Config files can contain invalid values at runtime.
  const { resolved, warnings } = resolveConfig({ subtitleSelection: { enabled: 'yes' } });
  assert.equal(resolved.subtitleSelection.enabled, false);
  assert.equal(warnings.length, 1);
  const field = buildConfigSettingsRegistry(resolved).find(
    (entry) => entry.configPath === 'subtitleSelection.enabled',
  );
  assert.equal(field?.category, 'behavior');
  assert.equal(field?.restartBehavior, 'hot-reload');
  const cleared = resolveConfig({
    subtitleSelection: { enabled: true },
    shortcuts: { openSubtitleSelection: null },
  }).resolved;
  assert.equal(resolveConfiguredShortcuts(cleared, DEFAULT_CONFIG).openSubtitleSelection, null);
});

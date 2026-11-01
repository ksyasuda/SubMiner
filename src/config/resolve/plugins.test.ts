import assert from 'node:assert/strict';
import test from 'node:test';
import { resolveConfig } from '../resolve';

test('resolveConfig keeps plugin entries and their settings', () => {
  const { resolved, warnings } = resolveConfig({
    plugins: {
      'ai-explain': { enabled: true, settings: { codexPath: 'codex', before: 2 } },
      dormant: {},
    },
  });

  assert.deepEqual(resolved.plugins, {
    'ai-explain': { enabled: true, settings: { codexPath: 'codex', before: 2 } },
    dormant: { enabled: false, settings: {} },
  });
  assert.deepEqual(warnings, []);
});

test('resolveConfig warns about malformed plugin entries and keeps them disabled', () => {
  const { resolved, warnings } = resolveConfig({
    plugins: {
      flag: { enabled: 'yes' as never, settings: [] as never },
      scalar: true as never,
    },
  });

  assert.deepEqual(resolved.plugins, { flag: { enabled: false, settings: {} } });
  assert.deepEqual(
    warnings.map((warning) => warning.path),
    ['plugins.flag.enabled', 'plugins.flag.settings', 'plugins.scalar'],
  );
});

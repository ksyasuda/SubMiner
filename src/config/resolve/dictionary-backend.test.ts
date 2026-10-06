import assert from 'node:assert/strict';
import test from 'node:test';
import { resolveConfig } from '../resolve';
import { buildConfigSettingsRegistry } from '../settings/registry';
import { createResolveContext } from './context';
import { applyCoreDomainConfig } from './core-domains';

test('dictionary backend defaults to Yomitan and accepts Hachidori', () => {
  assert.equal(resolveConfig({}).resolved.dictionaryBackend, 'yomitan');
  const { resolved, warnings } = resolveConfig({ dictionaryBackend: 'hachidori' });
  assert.equal(resolved.dictionaryBackend, 'hachidori');
  assert.deepEqual(warnings, []);
  const field = buildConfigSettingsRegistry(resolved).find(
    (entry) => entry.configPath === 'dictionaryBackend',
  );
  assert.equal(field?.restartBehavior, 'restart');
  assert.deepEqual(field?.enumValues, ['yomitan', 'hachidori']);
  assert.equal(field?.category, 'integrations');
});

test('unknown dictionary backend values warn and preserve the default', () => {
  for (const dictionaryBackend of ['unknown', '', null, true, {}]) {
    const { context, warnings } = createResolveContext({});
    context.src.dictionaryBackend = dictionaryBackend;
    applyCoreDomainConfig(context);
    assert.equal(context.resolved.dictionaryBackend, 'yomitan');
    assert.equal(warnings.length, 1);
    assert.equal(warnings[0]?.path, 'dictionaryBackend');
  }
});

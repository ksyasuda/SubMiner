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

test('Hachidori external import URL accepts HTTP origins and rejects invalid targets', () => {
  assert.equal(resolveConfig({}).resolved.hachidori.externalHostManagementUrl, '');
  const result = resolveConfig({
    hachidori: { externalHostManagementUrl: 'http://127.0.0.1:8780/' },
  });
  assert.equal(result.resolved.hachidori.externalHostManagementUrl, 'http://127.0.0.1:8780');
  assert.deepEqual(result.warnings, []);
  for (const value of [
    'file:///tmp/dict',
    'http://host/import',
    'http://user:password@host',
    true,
  ]) {
    const { context, warnings } = createResolveContext({});
    context.src.hachidori = { externalHostManagementUrl: value };
    applyCoreDomainConfig(context);
    assert.equal(context.resolved.hachidori.externalHostManagementUrl, '');
    assert.equal(warnings[0]?.path, 'hachidori.externalHostManagementUrl');
  }
});

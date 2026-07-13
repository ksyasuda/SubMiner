import assert from 'node:assert/strict';
import test from 'node:test';
import { resolveConfig } from '../resolve';

test('resolveConfig maps legacy Animetosho settings to TsukiHime', () => {
  const { resolved, warnings } = resolveConfig({
    animetosho: {
      apiBaseUrl: 'https://legacy.example/v1',
      maxSearchResults: 23,
    },
    shortcuts: {
      openAnimetosho: 'Ctrl+Alt+T',
    },
  });

  assert.equal(resolved.tsukihime.apiBaseUrl, 'https://legacy.example/v1');
  assert.equal(resolved.tsukihime.maxSearchResults, 23);
  assert.equal(resolved.shortcuts.openTsukihime, 'Ctrl+Alt+T');
  assert.deepEqual(warnings, []);
});

test('resolveConfig gives current TsukiHime settings precedence over legacy aliases', () => {
  const { resolved, warnings } = resolveConfig({
    animetosho: {
      apiBaseUrl: 'https://legacy.example/v1',
      maxSearchResults: 23,
    },
    tsukihime: {
      apiBaseUrl: 'https://current.example/v1',
      maxSearchResults: 7,
    },
    shortcuts: {
      openAnimetosho: 'Ctrl+Alt+T',
      openTsukihime: null,
    },
  });

  assert.equal(resolved.tsukihime.apiBaseUrl, 'https://current.example/v1');
  assert.equal(resolved.tsukihime.maxSearchResults, 7);
  assert.equal(resolved.shortcuts.openTsukihime, null);
  assert.deepEqual(warnings, []);
});

test('resolveConfig does not fall back when the current TsukiHime setting is invalid', () => {
  const { resolved, warnings } = resolveConfig({
    animetosho: {
      apiBaseUrl: 'https://legacy.example/v1',
      maxSearchResults: 23,
    },
    tsukihime: 'invalid' as never,
  });

  assert.equal(resolved.tsukihime.apiBaseUrl, 'https://api.tsukihime.org/v1');
  assert.equal(resolved.tsukihime.maxSearchResults, 10);
  assert.deepEqual(warnings, [
    {
      path: 'tsukihime',
      value: 'invalid',
      fallback: resolved.tsukihime,
      message: 'Expected object.',
    },
  ]);
});

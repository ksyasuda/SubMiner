import assert from 'node:assert/strict';
import test from 'node:test';

import {
  CONFIG_OPTION_REGISTRY,
  CONFIG_TEMPLATE_SECTIONS,
  DEFAULT_CONFIG,
  DEFAULT_KEYBINDINGS,
  RUNTIME_OPTION_REGISTRY,
} from '../definitions';
import { buildCoreConfigOptionRegistry } from './options-core';
import { buildImmersionConfigOptionRegistry } from './options-immersion';
import { buildIntegrationConfigOptionRegistry } from './options-integrations';
import { buildSubtitleConfigOptionRegistry } from './options-subtitle';

test('config option registry includes critical paths and has unique entries', () => {
  const paths = CONFIG_OPTION_REGISTRY.map((entry) => entry.path);

  for (const requiredPath of [
    'logging.level',
    'annotationWebsocket.enabled',
    'controller.enabled',
    'controller.scrollPixelsPerSecond',
    'startupWarmups.lowPowerMode',
    'subtitleStyle.enableJlpt',
    'subtitleStyle.autoPauseVideoOnYomitanPopup',
    'ankiConnect.enabled',
    'anilist.characterDictionary.enabled',
    'anilist.characterDictionary.collapsibleSections.description',
    'immersionTracking.enabled',
  ]) {
    assert.ok(paths.includes(requiredPath), `missing config path: ${requiredPath}`);
  }

  assert.equal(new Set(paths).size, paths.length);
});

test('config template sections include expected domains and unique keys', () => {
  const keys = CONFIG_TEMPLATE_SECTIONS.map((section) => section.key);
  const requiredKeys: (typeof keys)[number][] = [
    'websocket',
    'annotationWebsocket',
    'controller',
    'startupWarmups',
    'subtitleStyle',
    'ankiConnect',
    'immersionTracking',
  ];

  for (const requiredKey of requiredKeys) {
    assert.ok(keys.includes(requiredKey), `missing template section key: ${requiredKey}`);
  }

  assert.equal(new Set(keys).size, keys.length);
});

test('domain registry builders each contribute entries to composed registry', () => {
  const domainEntries = [
    buildCoreConfigOptionRegistry(DEFAULT_CONFIG),
    buildSubtitleConfigOptionRegistry(DEFAULT_CONFIG),
    buildIntegrationConfigOptionRegistry(DEFAULT_CONFIG, RUNTIME_OPTION_REGISTRY),
    buildImmersionConfigOptionRegistry(DEFAULT_CONFIG),
  ];
  const composedPaths = new Set(CONFIG_OPTION_REGISTRY.map((entry) => entry.path));

  for (const entries of domainEntries) {
    assert.ok(entries.length > 0);
    assert.ok(entries.some((entry) => composedPaths.has(entry.path)));
  }
});

test('default keybindings include primary and secondary subtitle track cycling on J keys', () => {
  const keybindingMap = new Map(
    DEFAULT_KEYBINDINGS.map((binding) => [binding.key, binding.command]),
  );
  assert.deepEqual(keybindingMap.get('KeyJ'), ['cycle', 'sid']);
  assert.deepEqual(keybindingMap.get('Shift+KeyJ'), ['cycle', 'secondary-sid']);
});

test('default keybindings include fullscreen on F', () => {
  const keybindingMap = new Map(
    DEFAULT_KEYBINDINGS.map((binding) => [binding.key, binding.command]),
  );
  assert.deepEqual(keybindingMap.get('KeyF'), ['cycle', 'fullscreen']);
});

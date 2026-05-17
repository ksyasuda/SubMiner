import assert from 'node:assert/strict';
import test from 'node:test';

import { ResolvedConfig } from '../../types/config';
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

function collectConfigLeafPaths(config: ResolvedConfig): string[] {
  const leaves: string[] = [];
  const visit = (value: unknown, prefix: string): void => {
    if (value === null || typeof value !== 'object' || Array.isArray(value)) {
      leaves.push(prefix);
      return;
    }
    const entries = Object.entries(value as Record<string, unknown>);
    if (entries.length === 0) {
      leaves.push(prefix);
      return;
    }
    for (const [key, child] of entries) {
      visit(child, prefix ? `${prefix}.${key}` : key);
    }
  };
  visit(config, '');
  return leaves;
}

// DEFAULT_CONFIG leaves that intentionally do not have a curated
// CONFIG_OPTION_REGISTRY entry. The generated config.example.jsonc still
// includes these paths, but their inline comments fall back to an auto-
// humanized key name instead of a written description.
//
// Current intentional gaps:
//   - subtitleStyle.*: thin wrappers around standard CSS properties; the
//     CSS reference is the canonical documentation surface.
//   - keybindings: an array of {key, command} objects, documented at the
//     section level via CONFIG_TEMPLATE_SECTIONS rather than per-leaf.
//
// New leaves added to DEFAULT_CONFIG should prefer a registry entry over
// an allowlist entry. Only allowlist a path when the registry is genuinely
// the wrong surface for it.
const UNDOCUMENTED_LEAVES: ReadonlySet<string> = new Set([
  'keybindings',
  'subtitleStyle.backdropFilter',
  'subtitleStyle.backgroundColor',
  'subtitleStyle.fontColor',
  'subtitleStyle.fontFamily',
  'subtitleStyle.fontKerning',
  'subtitleStyle.fontSize',
  'subtitleStyle.fontStyle',
  'subtitleStyle.fontWeight',
  'subtitleStyle.jlptColors.N1',
  'subtitleStyle.jlptColors.N2',
  'subtitleStyle.jlptColors.N3',
  'subtitleStyle.jlptColors.N4',
  'subtitleStyle.jlptColors.N5',
  'subtitleStyle.knownWordColor',
  'subtitleStyle.letterSpacing',
  'subtitleStyle.lineHeight',
  'subtitleStyle.nPlusOneColor',
  'subtitleStyle.secondary.backdropFilter',
  'subtitleStyle.secondary.backgroundColor',
  'subtitleStyle.secondary.fontColor',
  'subtitleStyle.secondary.fontFamily',
  'subtitleStyle.secondary.fontKerning',
  'subtitleStyle.secondary.fontSize',
  'subtitleStyle.secondary.fontStyle',
  'subtitleStyle.secondary.fontWeight',
  'subtitleStyle.secondary.letterSpacing',
  'subtitleStyle.secondary.lineHeight',
  'subtitleStyle.secondary.textRendering',
  'subtitleStyle.secondary.textShadow',
  'subtitleStyle.secondary.wordSpacing',
  'subtitleStyle.textRendering',
  'subtitleStyle.textShadow',
  'subtitleStyle.wordSpacing',
]);

test('config option registry includes critical paths and has unique entries', () => {
  const paths = CONFIG_OPTION_REGISTRY.map((entry) => entry.path);

  for (const requiredPath of [
    'logging.level',
    'annotationWebsocket.enabled',
    'controller.enabled',
    'controller.scrollPixelsPerSecond',
    'startupWarmups.lowPowerMode',
    'updates.channel',
    'youtube.primarySubLanguages',
    'subtitleStyle.enableJlpt',
    'subtitleStyle.autoPauseVideoOnYomitanPopup',
    'ankiConnect.enabled',
    'anilist.characterDictionary.enabled',
    'anilist.characterDictionary.collapsibleSections.description',
    'mpv.executablePath',
    'mpv.launchMode',
    'yomitan.externalProfilePath',
    'immersionTracking.enabled',
  ]) {
    assert.ok(paths.includes(requiredPath), `missing config path: ${requiredPath}`);
  }

  assert.equal(new Set(paths).size, paths.length);
});

test('every DEFAULT_CONFIG leaf is in CONFIG_OPTION_REGISTRY or UNDOCUMENTED_LEAVES', () => {
  const registryPaths = new Set(CONFIG_OPTION_REGISTRY.map((entry) => entry.path));
  const leaves = collectConfigLeafPaths(DEFAULT_CONFIG);

  const missing = leaves
    .filter((path) => !registryPaths.has(path) && !UNDOCUMENTED_LEAVES.has(path))
    .sort();
  assert.deepEqual(
    missing,
    [],
    `Add CONFIG_OPTION_REGISTRY entries (preferred) or add to UNDOCUMENTED_LEAVES allowlist: ${missing.join(', ')}`,
  );

  const stale = [...UNDOCUMENTED_LEAVES].filter((path) => registryPaths.has(path)).sort();
  assert.deepEqual(
    stale,
    [],
    `Remove from UNDOCUMENTED_LEAVES (now covered by CONFIG_OPTION_REGISTRY): ${stale.join(', ')}`,
  );

  const leafSet = new Set(leaves);
  const orphaned = [...UNDOCUMENTED_LEAVES].filter((path) => !leafSet.has(path)).sort();
  assert.deepEqual(
    orphaned,
    [],
    `Remove from UNDOCUMENTED_LEAVES (no longer a DEFAULT_CONFIG leaf): ${orphaned.join(', ')}`,
  );
});

test('config template sections include expected domains and unique keys', () => {
  const keys = CONFIG_TEMPLATE_SECTIONS.map((section) => section.key);
  const requiredKeys: (typeof keys)[number][] = [
    'websocket',
    'annotationWebsocket',
    'controller',
    'startupWarmups',
    'youtube',
    'subtitleStyle',
    'ankiConnect',
    'yomitan',
    'mpv',
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
  assert.deepEqual(keybindingMap.get('Ctrl+Alt+KeyC'), ['__youtube-picker-open']);
  assert.deepEqual(keybindingMap.get('Ctrl+Alt+KeyP'), ['__playlist-browser-open']);
});

test('default keybindings include fullscreen on F', () => {
  const keybindingMap = new Map(
    DEFAULT_KEYBINDINGS.map((binding) => [binding.key, binding.command]),
  );
  assert.deepEqual(keybindingMap.get('KeyF'), ['cycle', 'fullscreen']);
});

test('default keybindings include replay and next subtitle controls', () => {
  const keybindingMap = new Map(
    DEFAULT_KEYBINDINGS.map((binding) => [binding.key, binding.command]),
  );
  assert.deepEqual(keybindingMap.get('Ctrl+Shift+KeyH'), ['__replay-subtitle']);
  assert.deepEqual(keybindingMap.get('Ctrl+Shift+KeyL'), ['__play-next-subtitle']);
});

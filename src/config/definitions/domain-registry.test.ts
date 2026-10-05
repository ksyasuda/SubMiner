import assert from 'node:assert/strict';
import test from 'node:test';

import { ResolvedConfig } from '../../types/config';
import { CONFIG_OPTION_REGISTRY, CONFIG_TEMPLATE_SECTIONS, DEFAULT_CONFIG } from '../definitions';

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
// CONFIG_OPTION_REGISTRY entry.
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
  'anilist.characterDictionary.evictionPolicy',
  'anilist.characterDictionary.refreshTtlHours',
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
  'subtitleStyle.letterSpacing',
  'subtitleStyle.lineHeight',
  'subtitleStyle.paintOrder',
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
  'subtitleStyle.secondary.paintOrder',
  'subtitleStyle.secondary.textRendering',
  'subtitleStyle.secondary.textShadow',
  'subtitleStyle.secondary.WebkitTextStroke',
  'subtitleStyle.secondary.wordSpacing',
  'subtitleStyle.textRendering',
  'subtitleStyle.textShadow',
  'subtitleStyle.WebkitTextStroke',
  'subtitleStyle.wordSpacing',
]);

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

test('config option paths and template section keys are unique', () => {
  const paths = CONFIG_OPTION_REGISTRY.map((entry) => entry.path);
  assert.equal(new Set(paths).size, paths.length);

  const keys = CONFIG_TEMPLATE_SECTIONS.map((section) => section.key);
  assert.equal(new Set(keys).size, keys.length);
});

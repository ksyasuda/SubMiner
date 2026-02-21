import assert from 'node:assert/strict';
import test from 'node:test';
import { DEFAULT_CONFIG, deepCloneConfig } from '../definitions';
import { createWarningCollector } from '../warnings';
import { applyAnkiConnectResolution } from './anki-connect';
import type { ResolveContext } from './context';

function makeContext(ankiConnect: unknown): {
  context: ResolveContext;
  warnings: ReturnType<typeof createWarningCollector>['warnings'];
} {
  const { warnings, warn } = createWarningCollector();
  const resolved = deepCloneConfig(DEFAULT_CONFIG);
  const context = {
    src: { ankiConnect },
    resolved,
    warn,
  } as unknown as ResolveContext;

  return { context, warnings };
}

test('modern invalid nPlusOne.highlightEnabled warns modern key and does not fallback to legacy', () => {
  const { context, warnings } = makeContext({
    behavior: { nPlusOneHighlightEnabled: true },
    nPlusOne: { highlightEnabled: 'yes' },
  });

  applyAnkiConnectResolution(context);

  assert.equal(
    context.resolved.ankiConnect.nPlusOne.highlightEnabled,
    DEFAULT_CONFIG.ankiConnect.nPlusOne.highlightEnabled,
  );
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.nPlusOne.highlightEnabled'));
  assert.equal(
    warnings.some((warning) => warning.path === 'ankiConnect.behavior.nPlusOneHighlightEnabled'),
    false,
  );
});

test('normalizes ankiConnect tags by trimming and deduping', () => {
  const { context, warnings } = makeContext({
    tags: [' SubMiner ', 'Mining', 'SubMiner', '  Mining  '],
  });

  applyAnkiConnectResolution(context);

  assert.deepEqual(context.resolved.ankiConnect.tags, ['SubMiner', 'Mining']);
  assert.equal(
    warnings.some((warning) => warning.path === 'ankiConnect.tags'),
    false,
  );
});

test('warns and falls back for invalid nPlusOne.decks entries', () => {
  const { context, warnings } = makeContext({
    nPlusOne: { decks: ['Core Deck', 123] },
  });

  applyAnkiConnectResolution(context);

  assert.deepEqual(
    context.resolved.ankiConnect.nPlusOne.decks,
    DEFAULT_CONFIG.ankiConnect.nPlusOne.decks,
  );
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.nPlusOne.decks'));
});

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

test('modern invalid knownWords.highlightEnabled warns modern key and does not fallback to legacy', () => {
  const { context, warnings } = makeContext({
    nPlusOne: { highlightEnabled: true },
    knownWords: { highlightEnabled: 'yes' },
  });

  applyAnkiConnectResolution(context);

  assert.equal(
    context.resolved.ankiConnect.knownWords.highlightEnabled,
    DEFAULT_CONFIG.ankiConnect.knownWords.highlightEnabled,
  );
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.knownWords.highlightEnabled'));
  assert.equal(
    warnings.some((warning) => warning.path === 'ankiConnect.nPlusOne.highlightEnabled'),
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

test('accepts knownWords.decks object format with field arrays', () => {
  const { context, warnings } = makeContext({
    knownWords: { decks: { 'Core Deck': ['Word', 'Reading'], Mining: ['Expression'] } },
  });

  applyAnkiConnectResolution(context);

  assert.deepEqual(context.resolved.ankiConnect.knownWords.decks, {
    'Core Deck': ['Word', 'Reading'],
    Mining: ['Expression'],
  });
  assert.equal(
    warnings.some((warning) => warning.path === 'ankiConnect.knownWords.decks'),
    false,
  );
});

test('converts legacy knownWords.decks array to object with default fields', () => {
  const { context, warnings } = makeContext({
    knownWords: { decks: ['Core Deck'] },
  });

  applyAnkiConnectResolution(context);

  assert.deepEqual(context.resolved.ankiConnect.knownWords.decks, {
    'Core Deck': ['Expression', 'Word', 'Reading', 'Word Reading'],
  });
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.knownWords.decks'));
});

test('accepts valid proxy settings', () => {
  const { context, warnings } = makeContext({
    proxy: {
      enabled: true,
      host: '127.0.0.1',
      port: 9999,
      upstreamUrl: 'http://127.0.0.1:8765',
    },
  });

  applyAnkiConnectResolution(context);

  assert.equal(context.resolved.ankiConnect.proxy.enabled, true);
  assert.equal(context.resolved.ankiConnect.proxy.host, '127.0.0.1');
  assert.equal(context.resolved.ankiConnect.proxy.port, 9999);
  assert.equal(context.resolved.ankiConnect.proxy.upstreamUrl, 'http://127.0.0.1:8765');
  assert.equal(
    warnings.some((warning) => warning.path.startsWith('ankiConnect.proxy')),
    false,
  );
});

test('warns and falls back for invalid proxy settings', () => {
  const { context, warnings } = makeContext({
    proxy: {
      enabled: 'yes',
      host: '',
      port: -1,
      upstreamUrl: '',
    },
  });

  applyAnkiConnectResolution(context);

  assert.deepEqual(context.resolved.ankiConnect.proxy, DEFAULT_CONFIG.ankiConnect.proxy);
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.proxy.enabled'));
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.proxy.host'));
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.proxy.port'));
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.proxy.upstreamUrl'));
});

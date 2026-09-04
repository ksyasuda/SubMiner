import assert from 'node:assert/strict';
import test from 'node:test';
import { DEFAULT_CONFIG, deepCloneConfig } from '../definitions';
import { createWarningCollector } from '../warnings';
import { applyAnkiConnectResolution } from './anki-connect';
import { applyAnkiKnownWordsResolution } from './anki-connect/known-words';
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

test('media timing review is disabled by default and accepts a boolean override', () => {
  const defaultContext = makeContext({});
  applyAnkiConnectResolution(defaultContext.context);
  assert.equal(defaultContext.context.resolved.ankiConnect.media.reviewTiming, false);

  const enabledContext = makeContext({ media: { reviewTiming: true } });
  applyAnkiConnectResolution(enabledContext.context);
  assert.equal(enabledContext.context.resolved.ankiConnect.media.reviewTiming, true);
  assert.deepEqual(enabledContext.warnings, []);
});

test('modern media duration accepts zero as the disabled cap sentinel', () => {
  const disabledCap = makeContext({ media: { maxMediaDuration: 0 } });
  applyAnkiConnectResolution(disabledCap.context);
  assert.equal(disabledCap.context.resolved.ankiConnect.media.maxMediaDuration, 0);
  assert.deepEqual(disabledCap.warnings, []);

  const invalidCap = makeContext({ media: { maxMediaDuration: -1 } });
  applyAnkiConnectResolution(invalidCap.context);
  assert.equal(
    invalidCap.context.resolved.ankiConnect.media.maxMediaDuration,
    DEFAULT_CONFIG.ankiConnect.media.maxMediaDuration,
  );
  assert.ok(
    invalidCap.warnings.some((warning) => warning.path === 'ankiConnect.media.maxMediaDuration'),
  );
});

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

test('invalid modern known-words primitive values warn and keep defaults', () => {
  const { context, warnings } = makeContext({
    knownWords: {
      refreshMinutes: 'daily',
      matchMode: false,
    },
    nPlusOne: {
      minSentenceWords: 'three',
    },
  });

  applyAnkiConnectResolution(context);

  assert.equal(
    context.resolved.ankiConnect.knownWords.refreshMinutes,
    DEFAULT_CONFIG.ankiConnect.knownWords.refreshMinutes,
  );
  assert.equal(
    context.resolved.ankiConnect.knownWords.matchMode,
    DEFAULT_CONFIG.ankiConnect.knownWords.matchMode,
  );
  assert.equal(
    context.resolved.ankiConnect.nPlusOne.minSentenceWords,
    DEFAULT_CONFIG.ankiConnect.nPlusOne.minSentenceWords,
  );
  assert.deepEqual(
    warnings.map((warning) => warning.path),
    [
      'ankiConnect.knownWords.refreshMinutes',
      'ankiConnect.nPlusOne.minSentenceWords',
      'ankiConnect.knownWords.matchMode',
    ],
  );
});

test('invalid legacy known-words primitive values warn and keep defaults', () => {
  const { context, warnings } = makeContext({
    behavior: {
      nPlusOneHighlightEnabled: 'yes',
      nPlusOneRefreshMinutes: 'daily',
      nPlusOneMatchMode: false,
    },
  });

  applyAnkiConnectResolution(context);

  assert.equal(
    context.resolved.ankiConnect.knownWords.highlightEnabled,
    DEFAULT_CONFIG.ankiConnect.knownWords.highlightEnabled,
  );
  assert.equal(
    context.resolved.ankiConnect.knownWords.refreshMinutes,
    DEFAULT_CONFIG.ankiConnect.knownWords.refreshMinutes,
  );
  assert.equal(
    context.resolved.ankiConnect.knownWords.matchMode,
    DEFAULT_CONFIG.ankiConnect.knownWords.matchMode,
  );
  assert.deepEqual(
    warnings.map((warning) => warning.path),
    [
      'ankiConnect.behavior.nPlusOneHighlightEnabled',
      'ankiConnect.behavior.nPlusOneRefreshMinutes',
      'ankiConnect.behavior.nPlusOneMatchMode',
    ],
  );
});

test('known-words resolution can run independently from other Anki domains', () => {
  const { context, warnings } = makeContext({
    knownWords: { highlightEnabled: true },
    proxy: { port: -1 },
  });
  const ankiConnect = context.src.ankiConnect as Record<string, unknown>;

  applyAnkiKnownWordsResolution(context, ankiConnect, {});

  assert.equal(context.resolved.ankiConnect.knownWords.highlightEnabled, true);
  assert.equal(
    warnings.some((warning) => warning.path.startsWith('ankiConnect.proxy')),
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

test('accepts knownWords.addMinedWordsImmediately boolean override', () => {
  const { context, warnings } = makeContext({
    knownWords: { addMinedWordsImmediately: false },
  });

  applyAnkiConnectResolution(context);

  assert.equal(context.resolved.ankiConnect.knownWords.addMinedWordsImmediately, false);
  assert.equal(
    warnings.some((warning) => warning.path === 'ankiConnect.knownWords.addMinedWordsImmediately'),
    false,
  );
});

test('knownWords.highlightEnabled does not implicitly enable nPlusOne', () => {
  const { context } = makeContext({
    knownWords: { highlightEnabled: true },
  });

  applyAnkiConnectResolution(context);

  assert.equal(context.resolved.ankiConnect.knownWords.highlightEnabled, true);
  assert.equal(
    context.resolved.ankiConnect.nPlusOne.enabled,
    DEFAULT_CONFIG.ankiConnect.nPlusOne.enabled,
  );
});

test('explicit nPlusOne.enabled is respected regardless of highlightEnabled', () => {
  const { context } = makeContext({
    knownWords: { highlightEnabled: true },
    nPlusOne: { enabled: false },
  });

  applyAnkiConnectResolution(context);

  assert.equal(context.resolved.ankiConnect.knownWords.highlightEnabled, true);
  assert.equal(context.resolved.ankiConnect.nPlusOne.enabled, false);
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

test('accepts configured ankiConnect.fields.word override', () => {
  const { context, warnings } = makeContext({
    fields: {
      word: 'TargetWord',
    },
  });

  applyAnkiConnectResolution(context);

  assert.equal(context.resolved.ankiConnect.fields.word, 'TargetWord');
  assert.equal(
    warnings.some((warning) => warning.path === 'ankiConnect.fields.word'),
    false,
  );
});

test('accepts ankiConnect.media.syncAnimatedImageToWordAudio override', () => {
  const { context, warnings } = makeContext({
    media: {
      syncAnimatedImageToWordAudio: false,
    },
  });

  applyAnkiConnectResolution(context);

  assert.equal(context.resolved.ankiConnect.media.syncAnimatedImageToWordAudio, false);
  assert.equal(
    warnings.some((warning) => warning.path === 'ankiConnect.media.syncAnimatedImageToWordAudio'),
    false,
  );
});

test('invalid modern Anki subtrees warn and keep resolved defaults', () => {
  const { context, warnings } = makeContext({
    fields: { word: 7 },
    media: { generateAudio: 'yes' },
    behavior: { overwriteAudio: 'yes' },
    metadata: { pattern: false },
  });

  applyAnkiConnectResolution(context);

  assert.equal(context.resolved.ankiConnect.fields.word, DEFAULT_CONFIG.ankiConnect.fields.word);
  assert.equal(
    context.resolved.ankiConnect.media.generateAudio,
    DEFAULT_CONFIG.ankiConnect.media.generateAudio,
  );
  assert.equal(
    context.resolved.ankiConnect.behavior.overwriteAudio,
    DEFAULT_CONFIG.ankiConnect.behavior.overwriteAudio,
  );
  assert.equal(
    context.resolved.ankiConnect.metadata.pattern,
    DEFAULT_CONFIG.ankiConnect.metadata.pattern,
  );
  assert.deepEqual(
    warnings.map((warning) => warning.path),
    [
      'ankiConnect.fields.word',
      'ankiConnect.media.generateAudio',
      'ankiConnect.behavior.overwriteAudio',
      'ankiConnect.metadata.pattern',
    ],
  );
});

test('maps legacy ankiConnect.wordField to modern ankiConnect.fields.word', () => {
  const { context, warnings } = makeContext({
    wordField: 'TargetWordLegacy',
  });

  applyAnkiConnectResolution(context);

  assert.equal(context.resolved.ankiConnect.fields.word, 'TargetWordLegacy');
  assert.equal(
    warnings.some((warning) => warning.path === 'ankiConnect.wordField'),
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

test('resolves knownWords maturity settings from config', () => {
  const { context, warnings } = makeContext({
    knownWords: { highlightEnabled: true, maturityEnabled: true, matureThresholdDays: 30 },
  });

  applyAnkiConnectResolution(context);

  assert.equal(context.resolved.ankiConnect.knownWords.maturityEnabled, true);
  assert.equal(context.resolved.ankiConnect.knownWords.matureThresholdDays, 30);
  assert.deepEqual(warnings, []);
});

test('warns and falls back for invalid knownWords maturity settings', () => {
  const { context, warnings } = makeContext({
    knownWords: { maturityEnabled: 'yes', matureThresholdDays: 0 },
  });

  applyAnkiConnectResolution(context);

  assert.equal(
    context.resolved.ankiConnect.knownWords.maturityEnabled,
    DEFAULT_CONFIG.ankiConnect.knownWords.maturityEnabled,
  );
  assert.equal(
    context.resolved.ankiConnect.knownWords.matureThresholdDays,
    DEFAULT_CONFIG.ankiConnect.knownWords.matureThresholdDays,
  );
  assert.ok(warnings.some((warning) => warning.path === 'ankiConnect.knownWords.maturityEnabled'));
  assert.ok(
    warnings.some((warning) => warning.path === 'ankiConnect.knownWords.matureThresholdDays'),
  );
});

test('omitted knownWords maturity settings fall back to defaults', () => {
  const { context, warnings } = makeContext({
    knownWords: { highlightEnabled: true },
  });

  applyAnkiConnectResolution(context);

  assert.equal(
    context.resolved.ankiConnect.knownWords.maturityEnabled,
    DEFAULT_CONFIG.ankiConnect.knownWords.maturityEnabled,
  );
  assert.equal(
    context.resolved.ankiConnect.knownWords.matureThresholdDays,
    DEFAULT_CONFIG.ankiConnect.knownWords.matureThresholdDays,
  );
  assert.deepEqual(warnings, []);
});

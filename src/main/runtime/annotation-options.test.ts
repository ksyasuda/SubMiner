import assert from 'node:assert/strict';
import test from 'node:test';
import { DEFAULT_CONFIG } from '../../config/definitions';
import type { ResolvedConfig } from '../../types';
import type { RuntimeOptionValue } from '../../types/runtime-options';
import { createAnnotationOptions, type AnnotationRuntimeOptionId } from './annotation-options';

/** Every annotation off in config; `runtime` holds the options the user has cycled. */
function setup(runtime: Partial<Record<AnnotationRuntimeOptionId, RuntimeOptionValue>> = {}) {
  const config: ResolvedConfig = {
    ...DEFAULT_CONFIG,
    ankiConnect: {
      ...DEFAULT_CONFIG.ankiConnect,
      knownWords: { ...DEFAULT_CONFIG.ankiConnect.knownWords, highlightEnabled: false },
      nPlusOne: { ...DEFAULT_CONFIG.ankiConnect.nPlusOne, enabled: false },
    },
    subtitleStyle: {
      ...DEFAULT_CONFIG.subtitleStyle,
      enableJlpt: false,
      nameMatchEnabled: true,
      frequencyDictionary: { ...DEFAULT_CONFIG.subtitleStyle.frequencyDictionary, enabled: false },
    },
  };
  return createAnnotationOptions({
    getConfig: () => config,
    getRuntimeOptionValue: (id) => runtime[id],
    isCharacterDictionaryEnabled: () => false,
  });
}

test('MeCab starts only when some annotation, from config or a runtime option, needs tokens', () => {
  assert.equal(setup().shouldInitializeMecabForAnnotations(), false);
  assert.equal(
    setup({ 'subtitle.annotation.jlpt': true }).shouldInitializeMecabForAnnotations(),
    true,
  );
});

test('a runtime option overrides config only when it holds a boolean', () => {
  const options = setup({ 'subtitle.annotation.nPlusOne': true, 'subtitle.annotation.jlpt': 'on' });

  assert.equal(options.getRuntimeBooleanOption('subtitle.annotation.nPlusOne', false), true);
  assert.equal(options.getRuntimeBooleanOption('subtitle.annotation.jlpt', false), false);
  assert.equal(options.getRuntimeBooleanOption('subtitle.annotation.frequency', true), true);
});

test('the texthooker colors names only while the character dictionary is enabled', () => {
  const settings = setup({
    'subtitle.annotation.knownWords.highlightEnabled': true,
  }).getTexthookerBootstrapSettings();

  assert.equal(settings.enableKnownWordColoring, true);
  assert.equal(settings.enableNPlusOneColoring, false);
  assert.equal(settings.characterDictionaryEnabled, false);
  assert.equal(settings.enableNameMatchColoring, false);
});

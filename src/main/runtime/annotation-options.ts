import type { TexthookerBootstrapSettings } from '../../core/services/texthooker';
import type { ResolvedConfig } from '../../types';
import type { RuntimeOptionId, RuntimeOptionValue } from '../../types/runtime-options';

/** The subtitle annotation toggles the user can flip at runtime from the options palette. */
export type AnnotationRuntimeOptionId = Extract<
  RuntimeOptionId,
  | 'subtitle.annotation.knownWords.highlightEnabled'
  | 'subtitle.annotation.nPlusOne'
  | 'subtitle.annotation.jlpt'
  | 'subtitle.annotation.frequency'
>;

export interface AnnotationOptionsDeps {
  getConfig: () => ResolvedConfig;
  /** The runtime option's current value; undefined before the options manager exists. */
  getRuntimeOptionValue: (id: AnnotationRuntimeOptionId) => RuntimeOptionValue | undefined;
  isCharacterDictionaryEnabled: () => boolean;
}

/**
 * Resolves which subtitle annotations are on. A runtime option the user cycled wins over its
 * config value; the texthooker page and the MeCab startup decision both read through here.
 */
export function createAnnotationOptions(deps: AnnotationOptionsDeps) {
  function getRuntimeBooleanOption(id: AnnotationRuntimeOptionId, fallback: boolean): boolean {
    const value = deps.getRuntimeOptionValue(id);
    return typeof value === 'boolean' ? value : fallback;
  }

  /** MeCab is only worth starting when at least one annotation needs tokens. */
  function shouldInitializeMecabForAnnotations(): boolean {
    const config = deps.getConfig();
    const knownWordsEnabled = getRuntimeBooleanOption(
      'subtitle.annotation.knownWords.highlightEnabled',
      config.ankiConnect.knownWords.highlightEnabled,
    );
    const nPlusOneEnabled = getRuntimeBooleanOption(
      'subtitle.annotation.nPlusOne',
      config.ankiConnect.nPlusOne.enabled,
    );
    const jlptEnabled = getRuntimeBooleanOption(
      'subtitle.annotation.jlpt',
      config.subtitleStyle.enableJlpt,
    );
    const frequencyEnabled = getRuntimeBooleanOption(
      'subtitle.annotation.frequency',
      config.subtitleStyle.frequencyDictionary.enabled,
    );
    return knownWordsEnabled || nPlusOneEnabled || jlptEnabled || frequencyEnabled;
  }

  /** Annotation colors and toggles the texthooker page renders with. */
  function getTexthookerBootstrapSettings(): TexthookerBootstrapSettings {
    const config = deps.getConfig();
    const characterDictionaryEnabled =
      config.subtitleStyle.nameMatchEnabled && deps.isCharacterDictionaryEnabled();
    const knownWordColoringEnabled = getRuntimeBooleanOption(
      'subtitle.annotation.knownWords.highlightEnabled',
      config.ankiConnect.knownWords.highlightEnabled,
    );
    const nPlusOneColoringEnabled = getRuntimeBooleanOption(
      'subtitle.annotation.nPlusOne',
      config.ankiConnect.nPlusOne.enabled,
    );

    return {
      enableKnownWordColoring: knownWordColoringEnabled,
      enableNPlusOneColoring: nPlusOneColoringEnabled,
      enableNameMatchColoring: config.subtitleStyle.nameMatchEnabled && characterDictionaryEnabled,
      enableFrequencyColoring: getRuntimeBooleanOption(
        'subtitle.annotation.frequency',
        config.subtitleStyle.frequencyDictionary.enabled,
      ),
      enableJlptColoring: getRuntimeBooleanOption(
        'subtitle.annotation.jlpt',
        config.subtitleStyle.enableJlpt,
      ),
      characterDictionaryEnabled,
      knownWordColor: config.subtitleStyle.knownWordColor,
      nPlusOneColor: config.subtitleStyle.nPlusOneColor,
      nameMatchColor: config.subtitleStyle.nameMatchColor,
      hoverTokenColor: config.subtitleStyle.hoverTokenColor,
      hoverTokenBackgroundColor: config.subtitleStyle.hoverTokenBackgroundColor,
      jlptColors: config.subtitleStyle.jlptColors,
      frequencyDictionary: {
        singleColor: config.subtitleStyle.frequencyDictionary.singleColor,
        bandedColors: config.subtitleStyle.frequencyDictionary.bandedColors,
      },
    };
  }

  return {
    getRuntimeBooleanOption,
    shouldInitializeMecabForAnnotations,
    getTexthookerBootstrapSettings,
  };
}

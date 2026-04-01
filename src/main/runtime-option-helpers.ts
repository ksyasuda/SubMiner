import type { ResolvedConfig } from '../types';

export function getRuntimeBooleanOption(
  getOptionValue: (
    id:
      | 'subtitle.annotation.nPlusOne'
      | 'subtitle.annotation.jlpt'
      | 'subtitle.annotation.frequency',
  ) => unknown,
  id: 'subtitle.annotation.nPlusOne' | 'subtitle.annotation.jlpt' | 'subtitle.annotation.frequency',
  fallback: boolean,
): boolean {
  const value = getOptionValue(id);
  return typeof value === 'boolean' ? value : fallback;
}

export function shouldInitializeMecabForAnnotations(input: {
  getResolvedConfig: () => ResolvedConfig;
  getRuntimeBooleanOption: (
    id:
      | 'subtitle.annotation.nPlusOne'
      | 'subtitle.annotation.jlpt'
      | 'subtitle.annotation.frequency',
    fallback: boolean,
  ) => boolean;
}): boolean {
  const config = input.getResolvedConfig();
  const nPlusOneEnabled = input.getRuntimeBooleanOption(
    'subtitle.annotation.nPlusOne',
    config.ankiConnect.knownWords.highlightEnabled,
  );
  const jlptEnabled = input.getRuntimeBooleanOption(
    'subtitle.annotation.jlpt',
    config.subtitleStyle.enableJlpt,
  );
  const frequencyEnabled = input.getRuntimeBooleanOption(
    'subtitle.annotation.frequency',
    config.subtitleStyle.frequencyDictionary.enabled,
  );
  return nPlusOneEnabled || jlptEnabled || frequencyEnabled;
}

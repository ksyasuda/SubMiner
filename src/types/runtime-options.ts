export type RuntimeOptionId =
  | 'anki.autoUpdateNewCards'
  | 'subtitle.annotation.knownWords.highlightEnabled'
  | 'subtitle.annotation.knownWords.maturityEnabled'
  | 'subtitle.annotation.nPlusOne'
  | 'subtitle.annotation.jlpt'
  | 'subtitle.annotation.frequency'
  | 'anki.kikuFieldGrouping'
  | 'anki.senrenFieldGrouping'
  | 'anki.nPlusOneMatchMode';

export type RuntimeOptionScope = 'ankiConnect' | 'subtitle';

export type RuntimeOptionValueType = 'boolean' | 'enum';

export type RuntimeOptionValue = boolean | string;

export interface RuntimeOptionState {
  id: RuntimeOptionId;
  label: string;
  scope: RuntimeOptionScope;
  valueType: RuntimeOptionValueType;
  value: RuntimeOptionValue;
  allowedValues: RuntimeOptionValue[];
  requiresRestart: boolean;
}

export interface RuntimeOptionApplyResult {
  ok: boolean;
  option?: RuntimeOptionState;
  osdMessage?: string;
  requiresRestart?: boolean;
  error?: string;
}

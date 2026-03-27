import type { SubtitleCue } from '../core/services/subtitle-cue-parser';

export enum PartOfSpeech {
  noun = 'noun',
  verb = 'verb',
  i_adjective = 'i_adjective',
  na_adjective = 'na_adjective',
  particle = 'particle',
  bound_auxiliary = 'bound_auxiliary',
  symbol = 'symbol',
  other = 'other',
}

export interface Token {
  word: string;
  partOfSpeech: PartOfSpeech;
  pos1: string;
  pos2: string;
  pos3: string;
  pos4: string;
  inflectionType: string;
  inflectionForm: string;
  headword: string;
  katakanaReading: string;
  pronunciation: string;
}

export interface MergedToken {
  surface: string;
  reading: string;
  headword: string;
  startPos: number;
  endPos: number;
  partOfSpeech: PartOfSpeech;
  pos1?: string;
  pos2?: string;
  pos3?: string;
  isMerged: boolean;
  isKnown: boolean;
  isNPlusOneTarget: boolean;
  isNameMatch?: boolean;
  jlptLevel?: JlptLevel;
  frequencyRank?: number;
}

export type FrequencyDictionaryLookup = (term: string) => number | null;

export type JlptLevel = 'N1' | 'N2' | 'N3' | 'N4' | 'N5';

export interface SubtitlePosition {
  yPercent: number;
}

export interface SubtitleStyle {
  fontSize: number;
}

export type SecondarySubMode = 'hidden' | 'visible' | 'hover';

export interface SecondarySubConfig {
  secondarySubLanguages?: string[];
  autoLoadSecondarySub?: boolean;
  defaultMode?: SecondarySubMode;
}

export type NPlusOneMatchMode = 'headword' | 'surface';
export type FrequencyDictionaryMatchMode = 'headword' | 'surface';

export interface SubtitleStyleConfig {
  enableJlpt?: boolean;
  preserveLineBreaks?: boolean;
  autoPauseVideoOnHover?: boolean;
  autoPauseVideoOnYomitanPopup?: boolean;
  hoverTokenColor?: string;
  hoverTokenBackgroundColor?: string;
  nameMatchEnabled?: boolean;
  nameMatchColor?: string;
  fontFamily?: string;
  fontSize?: number;
  fontColor?: string;
  fontWeight?: string | number;
  fontStyle?: string;
  lineHeight?: string | number;
  letterSpacing?: string;
  wordSpacing?: string | number;
  fontKerning?: string;
  textRendering?: string;
  textShadow?: string;
  backdropFilter?: string;
  backgroundColor?: string;
  nPlusOneColor?: string;
  knownWordColor?: string;
  jlptColors?: {
    N1: string;
    N2: string;
    N3: string;
    N4: string;
    N5: string;
  };
  frequencyDictionary?: {
    enabled?: boolean;
    sourcePath?: string;
    topX?: number;
    mode?: FrequencyDictionaryMode;
    matchMode?: FrequencyDictionaryMatchMode;
    singleColor?: string;
    bandedColors?: [string, string, string, string, string];
  };
  secondary?: {
    fontFamily?: string;
    fontSize?: number;
    fontColor?: string;
    fontWeight?: string | number;
    fontStyle?: string;
    lineHeight?: string | number;
    letterSpacing?: string;
    wordSpacing?: string | number;
    fontKerning?: string;
    textRendering?: string;
    textShadow?: string;
    backdropFilter?: string;
    backgroundColor?: string;
  };
}

export interface TokenPos1ExclusionConfig {
  defaults?: string[];
  add?: string[];
  remove?: string[];
}

export interface ResolvedTokenPos1ExclusionConfig {
  defaults: string[];
  add: string[];
  remove: string[];
}

export interface TokenPos2ExclusionConfig {
  defaults?: string[];
  add?: string[];
  remove?: string[];
}

export interface ResolvedTokenPos2ExclusionConfig {
  defaults: string[];
  add: string[];
  remove: string[];
}

export type FrequencyDictionaryMode = 'single' | 'banded';

export type { SubtitleCue };

export type SubtitleSidebarLayout = 'overlay' | 'embedded';

export interface SubtitleSidebarConfig {
  enabled?: boolean;
  autoOpen?: boolean;
  layout?: SubtitleSidebarLayout;
  toggleKey?: string;
  pauseVideoOnHover?: boolean;
  autoScroll?: boolean;
  maxWidth?: number;
  opacity?: number;
  backgroundColor?: string;
  textColor?: string;
  fontFamily?: string;
  fontSize?: number;
  timestampColor?: string;
  activeLineColor?: string;
  activeLineBackgroundColor?: string;
  hoverLineBackgroundColor?: string;
}

export interface SubtitleData {
  text: string;
  tokens: MergedToken[] | null;
  startTime?: number | null;
  endTime?: number | null;
}

export interface SubtitleSidebarSnapshot {
  cues: SubtitleCue[];
  currentTimeSec?: number | null;
  currentSubtitle: {
    text: string;
    startTime: number | null;
    endTime: number | null;
  };
  config: Required<SubtitleSidebarConfig>;
}

export interface SubtitleHoverTokenPayload {
  tokenIndex: number | null;
}

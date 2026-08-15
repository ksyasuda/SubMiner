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
  /** Dictionary-form reading of headword (kana), when the parser provides it. */
  headwordReading?: string;
  startPos: number;
  endPos: number;
  partOfSpeech: PartOfSpeech;
  pos1?: string;
  pos2?: string;
  pos3?: string;
  isMerged: boolean;
  isKnown: boolean;
  /** Anki card maturity for a known token; unset when tier data is unavailable. */
  knownMaturity?: KnownWordMaturityTier;
  isNPlusOneTarget: boolean;
  /**
   * Text Yomitan had no dictionary entry for (e.g. ぅ～ elongation runs,
   * truncated inflections). Kept as a token so it stays hoverable, but
   * ignored by annotation, N+1, and vocabulary-stats logic.
   */
  isUnparsedRun?: boolean;
  isNameMatch?: boolean;
  characterImage?: CharacterNameImage;
  jlptLevel?: JlptLevel;
  frequencyRank?: number;
}

export interface CharacterNameImage {
  src: string;
  alt: string;
}

export type FrequencyDictionaryLookup = (term: string) => number | null;

export type JlptLevel = 'N1' | 'N2' | 'N3' | 'N4' | 'N5';

/** Anki card maturity tier for a known word, most mature card wins. */
export type KnownWordMaturityTier = 'new' | 'learning' | 'young' | 'mature';

export interface SubtitlePosition {
  yPercent: number;
}

export interface SubtitleStyle {
  fontSize: number;
}

export type SubtitleBarMode = 'hidden' | 'visible' | 'hover';
export type PrimarySubMode = SubtitleBarMode;
export type SecondarySubMode = SubtitleBarMode;

export interface SecondarySubConfig {
  secondarySubLanguages?: string[];
  autoLoadSecondarySub?: boolean;
  defaultMode?: SecondarySubMode;
}

export type NPlusOneMatchMode = 'headword' | 'surface';
export type FrequencyDictionaryMatchMode = 'headword' | 'surface';

export interface SubtitleStyleConfig {
  primaryDefaultMode?: PrimarySubMode;
  css?: Record<string, string>;
  enableJlpt?: boolean;
  preserveLineBreaks?: boolean;
  autoPauseVideoOnHover?: boolean;
  autoPauseVideoOnYomitanPopup?: boolean;
  primaryVisibleOnYomitanPopup?: boolean;
  hoverTokenColor?: string;
  hoverTokenBackgroundColor?: string;
  nameMatchEnabled?: boolean;
  nameMatchImagesEnabled?: boolean;
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
  paintOrder?: string;
  WebkitTextStroke?: string;
  backdropFilter?: string;
  backgroundColor?: string;
  nPlusOneColor?: string;
  knownWordColor?: string;
  knownWordMaturityColors?: {
    new: string;
    learning: string;
    young: string;
    mature: string;
  };
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
    css?: Record<string, string>;
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
    paintOrder?: string;
    WebkitTextStroke?: string;
    backdropFilter?: string;
    backgroundColor?: string;
  };
}

export type SubtitleRendererStyleConfig = SubtitleStyleConfig;

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
  css?: Record<string, string>;
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

export type ResolvedSubtitleSidebarConfig = Required<Omit<SubtitleSidebarConfig, 'css'>> & {
  css: Record<string, string>;
};

export type SubtitleSidebarSnapshotConfig = Required<Omit<SubtitleSidebarConfig, 'css'>> & {
  css?: Record<string, string>;
};

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
  config: SubtitleSidebarSnapshotConfig;
}

export interface SubtitleMiningContext {
  source: 'subtitle-sidebar' | 'overlay';
  text: string;
  startTime: number;
  endTime: number;
  capturedAtMs?: number;
}

export interface SubtitleHoverTokenPayload {
  tokenIndex: number | null;
}

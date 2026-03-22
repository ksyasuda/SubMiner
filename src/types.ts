/*
 * SubMiner - All-in-one sentence mining overlay
 * Copyright (C) 2024 sudacode
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

import type { SubtitleCue } from './core/services/subtitle-cue-parser';

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

export interface WindowGeometry {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface SubtitlePosition {
  yPercent: number;
}

export interface SubtitleStyle {
  fontSize: number;
}

export interface Keybinding {
  key: string;
  command: (string | number)[] | null;
}

export type SecondarySubMode = 'hidden' | 'visible' | 'hover';

export interface SecondarySubConfig {
  secondarySubLanguages?: string[];
  autoLoadSecondarySub?: boolean;
  defaultMode?: SecondarySubMode;
}

export type SubsyncMode = 'auto' | 'manual';

export interface SubsyncConfig {
  defaultMode?: SubsyncMode;
  alass_path?: string;
  ffsubsync_path?: string;
  ffmpeg_path?: string;
  replace?: boolean;
}

export interface StartupWarmupsConfig {
  lowPowerMode?: boolean;
  mecab?: boolean;
  yomitanExtension?: boolean;
  subtitleDictionaries?: boolean;
  jellyfinRemoteSession?: boolean;
}

export interface WebSocketConfig {
  enabled?: boolean | 'auto';
  port?: number;
}

export interface AnnotationWebSocketConfig {
  enabled?: boolean;
  port?: number;
}

export interface TexthookerConfig {
  launchAtStartup?: boolean;
  openBrowser?: boolean;
}

export interface NotificationOptions {
  body?: string;
  icon?: string;
}

export interface MpvClient {
  currentSubText: string;
  currentVideoPath: string;
  currentMediaTitle?: string | null;
  currentTimePos: number;
  currentSubStart: number;
  currentSubEnd: number;
  currentAudioStreamIndex: number | null;
  send(command: { command: unknown[]; request_id?: number }): boolean;
}

export interface KikuDuplicateCardInfo {
  noteId: number;
  expression: string;
  sentencePreview: string;
  hasAudio: boolean;
  hasImage: boolean;
  isOriginal: boolean;
}

export interface KikuFieldGroupingRequestData {
  original: KikuDuplicateCardInfo;
  duplicate: KikuDuplicateCardInfo;
}

export interface KikuFieldGroupingChoice {
  keepNoteId: number;
  deleteNoteId: number;
  deleteDuplicate: boolean;
  cancelled: boolean;
}

export interface KikuMergePreviewRequest {
  keepNoteId: number;
  deleteNoteId: number;
  deleteDuplicate: boolean;
}

export interface KikuMergePreviewResponse {
  ok: boolean;
  compact?: Record<string, unknown>;
  full?: Record<string, unknown>;
  error?: string;
}

export type RuntimeOptionId =
  | 'anki.autoUpdateNewCards'
  | 'subtitle.annotation.nPlusOne'
  | 'subtitle.annotation.jlpt'
  | 'subtitle.annotation.frequency'
  | 'anki.kikuFieldGrouping'
  | 'anki.nPlusOneMatchMode';

export type RuntimeOptionScope = 'ankiConnect' | 'subtitle';

export type RuntimeOptionValueType = 'boolean' | 'enum';

export type RuntimeOptionValue = boolean | string;

export type NPlusOneMatchMode = 'headword' | 'surface';
export type FrequencyDictionaryMatchMode = 'headword' | 'surface';

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

export interface AnkiConnectConfig {
  enabled?: boolean;
  url?: string;
  pollingRate?: number;
  proxy?: {
    enabled?: boolean;
    host?: string;
    port?: number;
    upstreamUrl?: string;
  };
  tags?: string[];
  fields?: {
    word?: string;
    audio?: string;
    image?: string;
    sentence?: string;
    miscInfo?: string;
    translation?: string;
  };
  ai?: boolean | AiFeatureConfig;
  media?: {
    generateAudio?: boolean;
    generateImage?: boolean;
    imageType?: 'static' | 'avif';
    imageFormat?: 'jpg' | 'png' | 'webp';
    imageQuality?: number;
    imageMaxWidth?: number;
    imageMaxHeight?: number;
    animatedFps?: number;
    animatedMaxWidth?: number;
    animatedMaxHeight?: number;
    animatedCrf?: number;
    syncAnimatedImageToWordAudio?: boolean;
    audioPadding?: number;
    fallbackDuration?: number;
    maxMediaDuration?: number;
  };
  knownWords?: {
    highlightEnabled?: boolean;
    refreshMinutes?: number;
    addMinedWordsImmediately?: boolean;
    matchMode?: NPlusOneMatchMode;
    decks?: Record<string, string[]>;
    color?: string;
  };
  nPlusOne?: {
    nPlusOne?: string;
    minSentenceWords?: number;
  };
  behavior?: {
    overwriteAudio?: boolean;
    overwriteImage?: boolean;
    mediaInsertMode?: 'append' | 'prepend';
    highlightWord?: boolean;
    notificationType?: 'osd' | 'system' | 'both' | 'none';
    autoUpdateNewCards?: boolean;
  };
  metadata?: {
    pattern?: string;
  };
  deck?: string;
  isLapis?: {
    enabled?: boolean;
    sentenceCardModel?: string;
  };
  isKiku?: {
    enabled?: boolean;
    fieldGrouping?: 'auto' | 'manual' | 'disabled';
    deleteDuplicateInAuto?: boolean;
  };
}

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

export type { SubtitleCue } from './core/services/subtitle-cue-parser';

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

export interface ShortcutsConfig {
  toggleVisibleOverlayGlobal?: string | null;
  copySubtitle?: string | null;
  copySubtitleMultiple?: string | null;
  updateLastCardFromClipboard?: string | null;
  triggerFieldGrouping?: string | null;
  triggerSubsync?: string | null;
  mineSentence?: string | null;
  mineSentenceMultiple?: string | null;
  multiCopyTimeoutMs?: number;
  toggleSecondarySub?: string | null;
  markAudioCard?: string | null;
  openRuntimeOptions?: string | null;
  openJimaku?: string | null;
}

export type ControllerButtonBinding =
  | 'none'
  | 'select'
  | 'buttonSouth'
  | 'buttonEast'
  | 'buttonNorth'
  | 'buttonWest'
  | 'leftShoulder'
  | 'rightShoulder'
  | 'leftStickPress'
  | 'rightStickPress'
  | 'leftTrigger'
  | 'rightTrigger';

export type ControllerAxisBinding = 'leftStickX' | 'leftStickY' | 'rightStickX' | 'rightStickY';
export type ControllerTriggerInputMode = 'auto' | 'digital' | 'analog';
export type ControllerAxisDirection = 'negative' | 'positive';
export type ControllerDpadFallback = 'none' | 'horizontal' | 'vertical';

export interface ControllerNoneBinding {
  kind: 'none';
}

export interface ControllerButtonInputBinding {
  kind: 'button';
  buttonIndex: number;
}

export interface ControllerAxisDirectionInputBinding {
  kind: 'axis';
  axisIndex: number;
  direction: ControllerAxisDirection;
}

export interface ControllerAxisInputBinding {
  kind: 'axis';
  axisIndex: number;
  dpadFallback?: ControllerDpadFallback;
}

export type ControllerDiscreteBindingConfig =
  | ControllerButtonBinding
  | ControllerNoneBinding
  | ControllerButtonInputBinding
  | ControllerAxisDirectionInputBinding;

export type ResolvedControllerDiscreteBinding =
  | ControllerNoneBinding
  | ControllerButtonInputBinding
  | ControllerAxisDirectionInputBinding;

export type ControllerAxisBindingConfig =
  | ControllerAxisBinding
  | ControllerNoneBinding
  | ControllerAxisInputBinding;

export type ResolvedControllerAxisBinding =
  | ControllerNoneBinding
  | {
      kind: 'axis';
      axisIndex: number;
      dpadFallback: ControllerDpadFallback;
    };

export interface ControllerBindingsConfig {
  toggleLookup?: ControllerDiscreteBindingConfig;
  closeLookup?: ControllerDiscreteBindingConfig;
  toggleKeyboardOnlyMode?: ControllerDiscreteBindingConfig;
  mineCard?: ControllerDiscreteBindingConfig;
  quitMpv?: ControllerDiscreteBindingConfig;
  previousAudio?: ControllerDiscreteBindingConfig;
  nextAudio?: ControllerDiscreteBindingConfig;
  playCurrentAudio?: ControllerDiscreteBindingConfig;
  toggleMpvPause?: ControllerDiscreteBindingConfig;
  leftStickHorizontal?: ControllerAxisBindingConfig;
  leftStickVertical?: ControllerAxisBindingConfig;
  rightStickHorizontal?: ControllerAxisBindingConfig;
  rightStickVertical?: ControllerAxisBindingConfig;
}

export interface ResolvedControllerBindingsConfig {
  toggleLookup?: ResolvedControllerDiscreteBinding;
  closeLookup?: ResolvedControllerDiscreteBinding;
  toggleKeyboardOnlyMode?: ResolvedControllerDiscreteBinding;
  mineCard?: ResolvedControllerDiscreteBinding;
  quitMpv?: ResolvedControllerDiscreteBinding;
  previousAudio?: ResolvedControllerDiscreteBinding;
  nextAudio?: ResolvedControllerDiscreteBinding;
  playCurrentAudio?: ResolvedControllerDiscreteBinding;
  toggleMpvPause?: ResolvedControllerDiscreteBinding;
  leftStickHorizontal?: ResolvedControllerAxisBinding;
  leftStickVertical?: ResolvedControllerAxisBinding;
  rightStickHorizontal?: ResolvedControllerAxisBinding;
  rightStickVertical?: ResolvedControllerAxisBinding;
}

export interface ControllerButtonIndicesConfig {
  select?: number;
  buttonSouth?: number;
  buttonEast?: number;
  buttonNorth?: number;
  buttonWest?: number;
  leftShoulder?: number;
  rightShoulder?: number;
  leftStickPress?: number;
  rightStickPress?: number;
  leftTrigger?: number;
  rightTrigger?: number;
}

export interface ControllerConfig {
  enabled?: boolean;
  preferredGamepadId?: string;
  preferredGamepadLabel?: string;
  smoothScroll?: boolean;
  scrollPixelsPerSecond?: number;
  horizontalJumpPixels?: number;
  stickDeadzone?: number;
  triggerInputMode?: ControllerTriggerInputMode;
  triggerDeadzone?: number;
  repeatDelayMs?: number;
  repeatIntervalMs?: number;
  buttonIndices?: ControllerButtonIndicesConfig;
  bindings?: ControllerBindingsConfig;
}

export interface ControllerPreferenceUpdate {
  preferredGamepadId: string;
  preferredGamepadLabel: string;
}

export type ControllerConfigUpdate = ControllerConfig;

export interface ControllerDeviceInfo {
  id: string;
  index: number;
  mapping: string;
  connected: boolean;
}

export interface ControllerButtonSnapshot {
  value: number;
  pressed: boolean;
  touched?: boolean;
}

export interface ControllerRuntimeSnapshot {
  connectedGamepads: ControllerDeviceInfo[];
  activeGamepadId: string | null;
  rawAxes: number[];
  rawButtons: ControllerButtonSnapshot[];
}

export type JimakuLanguagePreference = 'ja' | 'en' | 'none';

export interface JimakuConfig {
  apiKey?: string;
  apiKeyCommand?: string;
  apiBaseUrl?: string;
  languagePreference?: JimakuLanguagePreference;
  maxEntryResults?: number;
}

export type AnilistCharacterDictionaryEvictionPolicy = 'disable' | 'delete';
export type AnilistCharacterDictionaryProfileScope = 'all' | 'active';
export type AnilistCharacterDictionaryCollapsibleSectionKey =
  | 'description'
  | 'characterInformation'
  | 'voicedBy';

export interface AnilistCharacterDictionaryCollapsibleSectionsConfig {
  description?: boolean;
  characterInformation?: boolean;
  voicedBy?: boolean;
}

export interface AnilistCharacterDictionaryConfig {
  enabled?: boolean;
  refreshTtlHours?: number;
  maxLoaded?: number;
  evictionPolicy?: AnilistCharacterDictionaryEvictionPolicy;
  profileScope?: AnilistCharacterDictionaryProfileScope;
  collapsibleSections?: AnilistCharacterDictionaryCollapsibleSectionsConfig;
}

export interface AnilistConfig {
  enabled?: boolean;
  accessToken?: string;
  characterDictionary?: AnilistCharacterDictionaryConfig;
}

export interface YomitanConfig {
  externalProfilePath?: string;
}

export interface JellyfinConfig {
  enabled?: boolean;
  serverUrl?: string;
  username?: string;
  deviceId?: string;
  clientName?: string;
  clientVersion?: string;
  defaultLibraryId?: string;
  remoteControlEnabled?: boolean;
  remoteControlAutoConnect?: boolean;
  autoAnnounce?: boolean;
  remoteControlDeviceName?: string;
  pullPictures?: boolean;
  iconCacheDir?: string;
  directPlayPreferred?: boolean;
  directPlayContainers?: string[];
  transcodeVideoCodec?: string;
}

export interface DiscordPresenceConfig {
  enabled?: boolean;
  updateIntervalMs?: number;
  debounceMs?: number;
}

export interface AiFeatureConfig {
  enabled?: boolean;
  model?: string;
  systemPrompt?: string;
}

export interface AiConfig {
  enabled?: boolean;
  apiKey?: string;
  apiKeyCommand?: string;
  baseUrl?: string;
  model?: string;
  systemPrompt?: string;
  requestTimeoutMs?: number;
}

export interface YoutubeSubgenConfig {
  whisperBin?: string;
  whisperModel?: string;
  whisperVadModel?: string;
  whisperThreads?: number;
  fixWithAi?: boolean;
  ai?: AiFeatureConfig;
  primarySubLanguages?: string[];
}

export interface StatsConfig {
  toggleKey?: string;
  markWatchedKey?: string;
  serverPort?: number;
  autoStartServer?: boolean;
  autoOpenBrowser?: boolean;
}

export type ImmersionTrackingRetentionMode = 'preset' | 'advanced';
export type ImmersionTrackingRetentionPreset = 'minimal' | 'balanced' | 'deep-history';

export interface ImmersionTrackingConfig {
  enabled?: boolean;
  dbPath?: string;
  batchSize?: number;
  flushIntervalMs?: number;
  queueCap?: number;
  payloadCapBytes?: number;
  maintenanceIntervalMs?: number;
  retentionMode?: ImmersionTrackingRetentionMode;
  retentionPreset?: ImmersionTrackingRetentionPreset;
  retention?: {
    eventsDays?: number;
    telemetryDays?: number;
    sessionsDays?: number;
    dailyRollupsDays?: number;
    monthlyRollupsDays?: number;
    vacuumIntervalDays?: number;
  };
  lifetimeSummaries?: {
    global?: boolean;
    anime?: boolean;
    media?: boolean;
  };
}

export interface Config {
  subtitlePosition?: SubtitlePosition;
  keybindings?: Keybinding[];
  websocket?: WebSocketConfig;
  annotationWebsocket?: AnnotationWebSocketConfig;
  texthooker?: TexthookerConfig;
  controller?: ControllerConfig;
  ankiConnect?: AnkiConnectConfig;
  shortcuts?: ShortcutsConfig;
  secondarySub?: SecondarySubConfig;
  subsync?: SubsyncConfig;
  startupWarmups?: StartupWarmupsConfig;
  subtitleStyle?: SubtitleStyleConfig;
  subtitleSidebar?: SubtitleSidebarConfig;
  auto_start_overlay?: boolean;
  jimaku?: JimakuConfig;
  anilist?: AnilistConfig;
  yomitan?: YomitanConfig;
  jellyfin?: JellyfinConfig;
  discordPresence?: DiscordPresenceConfig;
  ai?: AiConfig;
  youtubeSubgen?: YoutubeSubgenConfig;
  immersionTracking?: ImmersionTrackingConfig;
  stats?: StatsConfig;
  logging?: {
    level?: 'debug' | 'info' | 'warn' | 'error';
  };
}

export type RawConfig = Config;

export interface ResolvedConfig {
  subtitlePosition: SubtitlePosition;
  keybindings: Keybinding[];
  websocket: Required<WebSocketConfig>;
  annotationWebsocket: Required<AnnotationWebSocketConfig>;
  texthooker: Required<TexthookerConfig>;
  controller: {
    enabled: boolean;
    preferredGamepadId: string;
    preferredGamepadLabel: string;
    smoothScroll: boolean;
    scrollPixelsPerSecond: number;
    horizontalJumpPixels: number;
    stickDeadzone: number;
    triggerInputMode: ControllerTriggerInputMode;
    triggerDeadzone: number;
    repeatDelayMs: number;
    repeatIntervalMs: number;
    buttonIndices: Required<ControllerButtonIndicesConfig>;
    bindings: Required<ResolvedControllerBindingsConfig>;
  };
  ankiConnect: AnkiConnectConfig & {
    enabled: boolean;
    url: string;
    pollingRate: number;
    proxy: {
      enabled: boolean;
      host: string;
      port: number;
      upstreamUrl: string;
    };
    tags: string[];
    fields: {
      word: string;
      audio: string;
      image: string;
      sentence: string;
      miscInfo: string;
      translation: string;
    };
    ai: AiFeatureConfig & {
      enabled: boolean;
    };
    media: {
      generateAudio: boolean;
      generateImage: boolean;
      imageType: 'static' | 'avif';
      imageFormat: 'jpg' | 'png' | 'webp';
      imageQuality: number;
      imageMaxWidth?: number;
      imageMaxHeight?: number;
      animatedFps: number;
      animatedMaxWidth: number;
      animatedMaxHeight?: number;
      animatedCrf: number;
      syncAnimatedImageToWordAudio: boolean;
      audioPadding: number;
      fallbackDuration: number;
      maxMediaDuration: number;
    };
    knownWords: {
      highlightEnabled: boolean;
      refreshMinutes: number;
      addMinedWordsImmediately: boolean;
      matchMode: NPlusOneMatchMode;
      decks: Record<string, string[]>;
      color: string;
    };
    nPlusOne: {
      nPlusOne: string;
      minSentenceWords: number;
    };
    behavior: {
      overwriteAudio: boolean;
      overwriteImage: boolean;
      mediaInsertMode: 'append' | 'prepend';
      highlightWord: boolean;
      notificationType: 'osd' | 'system' | 'both' | 'none';
      autoUpdateNewCards: boolean;
    };
    metadata: {
      pattern: string;
    };
    isLapis: {
      enabled: boolean;
      sentenceCardModel: string;
    };
    isKiku: {
      enabled: boolean;
      fieldGrouping: 'auto' | 'manual' | 'disabled';
      deleteDuplicateInAuto: boolean;
    };
  };
  shortcuts: Required<ShortcutsConfig>;
  secondarySub: Required<SecondarySubConfig>;
  subsync: Required<SubsyncConfig>;
  startupWarmups: {
    lowPowerMode: boolean;
    mecab: boolean;
    yomitanExtension: boolean;
    subtitleDictionaries: boolean;
    jellyfinRemoteSession: boolean;
  };
  subtitleStyle: Required<Omit<SubtitleStyleConfig, 'secondary' | 'frequencyDictionary'>> & {
    secondary: Required<NonNullable<SubtitleStyleConfig['secondary']>>;
    frequencyDictionary: {
      enabled: boolean;
      sourcePath: string;
      topX: number;
      mode: FrequencyDictionaryMode;
      matchMode: FrequencyDictionaryMatchMode;
      singleColor: string;
      bandedColors: [string, string, string, string, string];
    };
  };
  subtitleSidebar: Required<SubtitleSidebarConfig>;
  auto_start_overlay: boolean;
  jimaku: JimakuConfig & {
    apiBaseUrl: string;
    languagePreference: JimakuLanguagePreference;
    maxEntryResults: number;
  };
  anilist: {
    enabled: boolean;
    accessToken: string;
    characterDictionary: {
      enabled: boolean;
      refreshTtlHours: number;
      maxLoaded: number;
      evictionPolicy: AnilistCharacterDictionaryEvictionPolicy;
      profileScope: AnilistCharacterDictionaryProfileScope;
      collapsibleSections: Required<AnilistCharacterDictionaryCollapsibleSectionsConfig>;
    };
  };
  yomitan: {
    externalProfilePath: string;
  };
  jellyfin: {
    enabled: boolean;
    serverUrl: string;
    username: string;
    deviceId: string;
    clientName: string;
    clientVersion: string;
    defaultLibraryId: string;
    remoteControlEnabled: boolean;
    remoteControlAutoConnect: boolean;
    autoAnnounce: boolean;
    remoteControlDeviceName: string;
    pullPictures: boolean;
    iconCacheDir: string;
    directPlayPreferred: boolean;
    directPlayContainers: string[];
    transcodeVideoCodec: string;
  };
  discordPresence: {
    enabled: boolean;
    updateIntervalMs: number;
    debounceMs: number;
  };
  ai: AiConfig & {
    enabled: boolean;
    apiKey: string;
    apiKeyCommand: string;
    baseUrl: string;
    model: string;
    systemPrompt: string;
    requestTimeoutMs: number;
  };
  youtubeSubgen: YoutubeSubgenConfig & {
    whisperBin: string;
    whisperModel: string;
    whisperVadModel: string;
    whisperThreads: number;
    fixWithAi: boolean;
    ai: AiFeatureConfig;
    primarySubLanguages: string[];
  };
  immersionTracking: {
    enabled: boolean;
    dbPath?: string;
    batchSize: number;
    flushIntervalMs: number;
    queueCap: number;
    payloadCapBytes: number;
    maintenanceIntervalMs: number;
    retentionMode: ImmersionTrackingRetentionMode;
    retentionPreset: ImmersionTrackingRetentionPreset;
    retention: {
      eventsDays: number;
      telemetryDays: number;
      sessionsDays: number;
      dailyRollupsDays: number;
      monthlyRollupsDays: number;
      vacuumIntervalDays: number;
    };
    lifetimeSummaries: {
      global: boolean;
      anime: boolean;
      media: boolean;
    };
  };
  stats: {
    toggleKey: string;
    markWatchedKey: string;
    serverPort: number;
    autoStartServer: boolean;
    autoOpenBrowser: boolean;
  };
  logging: {
    level: 'debug' | 'info' | 'warn' | 'error';
  };
}

export interface ConfigValidationWarning {
  path: string;
  value: unknown;
  fallback: unknown;
  message: string;
}

export interface SubsyncSourceTrack {
  id: number;
  label: string;
}

export interface SubsyncManualPayload {
  sourceTracks: SubsyncSourceTrack[];
}

export interface SubsyncManualRunRequest {
  engine: 'alass' | 'ffsubsync';
  sourceTrackId?: number | null;
}

export interface SubsyncResult {
  ok: boolean;
  message: string;
}

export interface ClipboardAppendResult {
  ok: boolean;
  message: string;
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

export interface MpvSubtitleRenderMetrics {
  subPos: number;
  subFontSize: number;
  subScale: number;
  subMarginY: number;
  subMarginX: number;
  subFont: string;
  subSpacing: number;
  subBold: boolean;
  subItalic: boolean;
  subBorderSize: number;
  subShadowOffset: number;
  subAssOverride: string;
  subScaleByWindow: boolean;
  subUseMargins: boolean;
  osdHeight: number;
  osdDimensions: {
    w: number;
    h: number;
    ml: number;
    mr: number;
    mt: number;
    mb: number;
  } | null;
}

export type OverlayLayer = 'visible';

export interface OverlayContentRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface OverlayContentMeasurement {
  layer: OverlayLayer;
  measuredAtMs: number;
  viewport: {
    width: number;
    height: number;
  };
  contentRect: OverlayContentRect | null;
}

export interface MecabStatus {
  available: boolean;
  enabled: boolean;
  path: string | null;
}

export type JimakuConfidence = 'high' | 'medium' | 'low';

export interface JimakuMediaInfo {
  title: string;
  season: number | null;
  episode: number | null;
  confidence: JimakuConfidence;
  filename: string;
  rawTitle: string;
}

export interface JimakuSearchQuery {
  query: string;
}

export interface JimakuEntryFlags {
  anime?: boolean;
  movie?: boolean;
  adult?: boolean;
  external?: boolean;
  unverified?: boolean;
}

export interface JimakuEntry {
  id: number;
  name: string;
  english_name?: string | null;
  japanese_name?: string | null;
  flags?: JimakuEntryFlags;
  last_modified?: string;
}

export interface JimakuFilesQuery {
  entryId: number;
  episode?: number | null;
}

export interface JimakuFileEntry {
  name: string;
  url: string;
  size: number;
  last_modified: string;
}

export interface JimakuDownloadQuery {
  entryId: number;
  url: string;
  name: string;
}

export interface JimakuApiError {
  error: string;
  code?: number;
  retryAfter?: number;
}

export type JimakuApiResponse<T> = { ok: true; data: T } | { ok: false; error: JimakuApiError };

export type JimakuDownloadResult =
  | { ok: true; path: string }
  | { ok: false; error: JimakuApiError };

export interface ConfigHotReloadPayload {
  keybindings: Keybinding[];
  subtitleStyle: SubtitleStyleConfig | null;
  subtitleSidebar: Required<SubtitleSidebarConfig>;
  secondarySubMode: SecondarySubMode;
}

export type ResolvedControllerConfig = ResolvedConfig['controller'];

export interface SubtitleHoverTokenPayload {
  tokenIndex: number | null;
}

export interface ElectronAPI {
  getOverlayLayer: () => 'visible' | 'modal' | null;
  onSubtitle: (callback: (data: SubtitleData) => void) => void;
  onVisibility: (callback: (visible: boolean) => void) => void;
  onSubtitlePosition: (callback: (position: SubtitlePosition | null) => void) => void;
  getOverlayVisibility: () => Promise<boolean>;
  getCurrentSubtitle: () => Promise<SubtitleData>;
  getCurrentSubtitleRaw: () => Promise<string>;
  getCurrentSubtitleAss: () => Promise<string>;
  getSubtitleSidebarSnapshot: () => Promise<SubtitleSidebarSnapshot>;
  getPlaybackPaused: () => Promise<boolean | null>;
  onSubtitleAss: (callback: (assText: string) => void) => void;
  setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => void;
  openYomitanSettings: () => void;
  recordYomitanLookup: () => void;
  getSubtitlePosition: () => Promise<SubtitlePosition | null>;
  saveSubtitlePosition: (position: SubtitlePosition) => void;
  getMecabStatus: () => Promise<MecabStatus>;
  setMecabEnabled: (enabled: boolean) => void;
  sendMpvCommand: (command: (string | number)[]) => void;
  getKeybindings: () => Promise<Keybinding[]>;
  getConfiguredShortcuts: () => Promise<Required<ShortcutsConfig>>;
  getStatsToggleKey: () => Promise<string>;
  getMarkWatchedKey: () => Promise<string>;
  markActiveVideoWatched: () => Promise<boolean>;
  getControllerConfig: () => Promise<ResolvedControllerConfig>;
  saveControllerConfig: (update: ControllerConfigUpdate) => Promise<void>;
  saveControllerPreference: (update: ControllerPreferenceUpdate) => Promise<void>;
  getJimakuMediaInfo: () => Promise<JimakuMediaInfo>;
  jimakuSearchEntries: (query: JimakuSearchQuery) => Promise<JimakuApiResponse<JimakuEntry[]>>;
  jimakuListFiles: (query: JimakuFilesQuery) => Promise<JimakuApiResponse<JimakuFileEntry[]>>;
  jimakuDownloadFile: (query: JimakuDownloadQuery) => Promise<JimakuDownloadResult>;
  quitApp: () => void;
  toggleDevTools: () => void;
  toggleOverlay: () => void;
  toggleStatsOverlay: () => void;
  getAnkiConnectStatus: () => Promise<boolean>;
  setAnkiConnectEnabled: (enabled: boolean) => void;
  clearAnkiConnectHistory: () => void;
  onSecondarySub: (callback: (text: string) => void) => void;
  onSecondarySubMode: (callback: (mode: SecondarySubMode) => void) => void;
  getSecondarySubMode: () => Promise<SecondarySubMode>;
  getCurrentSecondarySub: () => Promise<string>;
  focusMainWindow: () => Promise<void>;
  getSubtitleStyle: () => Promise<SubtitleStyleConfig | null>;
  onSubsyncManualOpen: (callback: (payload: SubsyncManualPayload) => void) => void;
  runSubsyncManual: (request: SubsyncManualRunRequest) => Promise<SubsyncResult>;
  onKikuFieldGroupingRequest: (callback: (data: KikuFieldGroupingRequestData) => void) => void;
  kikuBuildMergePreview: (request: KikuMergePreviewRequest) => Promise<KikuMergePreviewResponse>;
  kikuFieldGroupingRespond: (choice: KikuFieldGroupingChoice) => void;
  getRuntimeOptions: () => Promise<RuntimeOptionState[]>;
  setRuntimeOptionValue: (
    id: RuntimeOptionId,
    value: RuntimeOptionValue,
  ) => Promise<RuntimeOptionApplyResult>;
  cycleRuntimeOption: (id: RuntimeOptionId, direction: 1 | -1) => Promise<RuntimeOptionApplyResult>;
  onRuntimeOptionsChanged: (callback: (options: RuntimeOptionState[]) => void) => void;
  onOpenRuntimeOptions: (callback: () => void) => void;
  onOpenJimaku: (callback: () => void) => void;
  onKeyboardModeToggleRequested: (callback: () => void) => void;
  onLookupWindowToggleRequested: (callback: () => void) => void;
  appendClipboardVideoToQueue: () => Promise<ClipboardAppendResult>;
  notifyOverlayModalClosed: (
    modal:
      | 'runtime-options'
      | 'subsync'
      | 'jimaku'
      | 'kiku'
      | 'controller-select'
      | 'controller-debug'
      | 'subtitle-sidebar',
  ) => void;
  notifyOverlayModalOpened: (
    modal:
      | 'runtime-options'
      | 'subsync'
      | 'jimaku'
      | 'kiku'
      | 'controller-select'
      | 'controller-debug'
      | 'subtitle-sidebar',
  ) => void;
  reportOverlayContentBounds: (measurement: OverlayContentMeasurement) => void;
  onConfigHotReload: (callback: (payload: ConfigHotReloadPayload) => void) => void;
}

declare global {
  interface Window {
    electronAPI: ElectronAPI;
  }
}

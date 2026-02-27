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
}

export interface WebSocketConfig {
  enabled?: boolean | 'auto';
  port?: number;
}

export interface TexthookerConfig {
  openBrowser?: boolean;
}

export interface NotificationOptions {
  body?: string;
  icon?: string;
}

export interface MpvClient {
  currentSubText: string;
  currentVideoPath: string;
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
  | 'anki.kikuFieldGrouping'
  | 'anki.nPlusOneMatchMode';

export type RuntimeOptionScope = 'ankiConnect';

export type RuntimeOptionValueType = 'boolean' | 'enum';

export type RuntimeOptionValue = boolean | string;

export type NPlusOneMatchMode = 'headword' | 'surface';

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
  tags?: string[];
  fields?: {
    audio?: string;
    image?: string;
    sentence?: string;
    miscInfo?: string;
    translation?: string;
  };
  ai?: {
    enabled?: boolean;
    alwaysUseAiTranslation?: boolean;
    apiKey?: string;
    model?: string;
    baseUrl?: string;
    targetLanguage?: string;
    systemPrompt?: string;
  };
  openRouter?: {
    enabled?: boolean;
    alwaysUseAiTranslation?: boolean;
    apiKey?: string;
    model?: string;
    baseUrl?: string;
    targetLanguage?: string;
    systemPrompt?: string;
  };
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
    audioPadding?: number;
    fallbackDuration?: number;
    maxMediaDuration?: number;
  };
  nPlusOne?: {
    highlightEnabled?: boolean;
    refreshMinutes?: number;
    matchMode?: NPlusOneMatchMode;
    decks?: string[];
    nPlusOne?: string;
    knownWord?: string;
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
  hoverTokenColor?: string;
  hoverTokenBackgroundColor?: string;
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

export type FrequencyDictionaryMode = 'single' | 'banded';

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

export type JimakuLanguagePreference = 'ja' | 'en' | 'none';

export interface JimakuConfig {
  apiKey?: string;
  apiKeyCommand?: string;
  apiBaseUrl?: string;
  languagePreference?: JimakuLanguagePreference;
  maxEntryResults?: number;
}

export interface AnilistConfig {
  enabled?: boolean;
  accessToken?: string;
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

export type YoutubeSubgenMode = 'automatic' | 'preprocess' | 'off';

export interface YoutubeSubgenConfig {
  mode?: YoutubeSubgenMode;
  whisperBin?: string;
  whisperModel?: string;
  primarySubLanguages?: string[];
}

export interface ImmersionTrackingConfig {
  enabled?: boolean;
  dbPath?: string;
  batchSize?: number;
  flushIntervalMs?: number;
  queueCap?: number;
  payloadCapBytes?: number;
  maintenanceIntervalMs?: number;
  retention?: {
    eventsDays?: number;
    telemetryDays?: number;
    dailyRollupsDays?: number;
    monthlyRollupsDays?: number;
    vacuumIntervalDays?: number;
  };
}

export interface Config {
  subtitlePosition?: SubtitlePosition;
  keybindings?: Keybinding[];
  websocket?: WebSocketConfig;
  texthooker?: TexthookerConfig;
  ankiConnect?: AnkiConnectConfig;
  shortcuts?: ShortcutsConfig;
  secondarySub?: SecondarySubConfig;
  subsync?: SubsyncConfig;
  subtitleStyle?: SubtitleStyleConfig;
  auto_start_overlay?: boolean;
  bind_visible_overlay_to_mpv_sub_visibility?: boolean;
  jimaku?: JimakuConfig;
  anilist?: AnilistConfig;
  jellyfin?: JellyfinConfig;
  discordPresence?: DiscordPresenceConfig;
  youtubeSubgen?: YoutubeSubgenConfig;
  immersionTracking?: ImmersionTrackingConfig;
  logging?: {
    level?: 'debug' | 'info' | 'warn' | 'error';
  };
}

export type RawConfig = Config;

export interface ResolvedConfig {
  subtitlePosition: SubtitlePosition;
  keybindings: Keybinding[];
  websocket: Required<WebSocketConfig>;
  texthooker: Required<TexthookerConfig>;
  ankiConnect: AnkiConnectConfig & {
    enabled: boolean;
    url: string;
    pollingRate: number;
    tags: string[];
    fields: {
      audio: string;
      image: string;
      sentence: string;
      miscInfo: string;
      translation: string;
    };
    ai: {
      enabled: boolean;
      alwaysUseAiTranslation: boolean;
      apiKey: string;
      model: string;
      baseUrl: string;
      targetLanguage: string;
      systemPrompt: string;
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
      audioPadding: number;
      fallbackDuration: number;
      maxMediaDuration: number;
    };
    nPlusOne: {
      highlightEnabled: boolean;
      refreshMinutes: number;
      matchMode: NPlusOneMatchMode;
      decks: string[];
      nPlusOne: string;
      knownWord: string;
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
  subtitleStyle: Required<Omit<SubtitleStyleConfig, 'secondary' | 'frequencyDictionary'>> & {
    secondary: Required<NonNullable<SubtitleStyleConfig['secondary']>>;
    frequencyDictionary: {
      enabled: boolean;
      sourcePath: string;
      topX: number;
      mode: FrequencyDictionaryMode;
      singleColor: string;
      bandedColors: [string, string, string, string, string];
    };
  };
  auto_start_overlay: boolean;
  bind_visible_overlay_to_mpv_sub_visibility: boolean;
  jimaku: JimakuConfig & {
    apiBaseUrl: string;
    languagePreference: JimakuLanguagePreference;
    maxEntryResults: number;
  };
  anilist: {
    enabled: boolean;
    accessToken: string;
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
  youtubeSubgen: YoutubeSubgenConfig & {
    mode: YoutubeSubgenMode;
    whisperBin: string;
    whisperModel: string;
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
    retention: {
      eventsDays: number;
      telemetryDays: number;
      dailyRollupsDays: number;
      monthlyRollupsDays: number;
      vacuumIntervalDays: number;
    };
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
  secondarySubMode: SecondarySubMode;
}

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
  onSubtitleAss: (callback: (assText: string) => void) => void;
  setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => void;
  openYomitanSettings: () => void;
  getSubtitlePosition: () => Promise<SubtitlePosition | null>;
  saveSubtitlePosition: (position: SubtitlePosition) => void;
  getMecabStatus: () => Promise<MecabStatus>;
  setMecabEnabled: (enabled: boolean) => void;
  sendMpvCommand: (command: (string | number)[]) => void;
  getKeybindings: () => Promise<Keybinding[]>;
  getConfiguredShortcuts: () => Promise<Required<ShortcutsConfig>>;
  getJimakuMediaInfo: () => Promise<JimakuMediaInfo>;
  jimakuSearchEntries: (query: JimakuSearchQuery) => Promise<JimakuApiResponse<JimakuEntry[]>>;
  jimakuListFiles: (query: JimakuFilesQuery) => Promise<JimakuApiResponse<JimakuFileEntry[]>>;
  jimakuDownloadFile: (query: JimakuDownloadQuery) => Promise<JimakuDownloadResult>;
  quitApp: () => void;
  toggleDevTools: () => void;
  toggleOverlay: () => void;
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
  appendClipboardVideoToQueue: () => Promise<ClipboardAppendResult>;
  notifyOverlayModalClosed: (modal: 'runtime-options' | 'subsync' | 'jimaku' | 'kiku') => void;
  notifyOverlayModalOpened: (modal: 'runtime-options' | 'subsync' | 'jimaku' | 'kiku') => void;
  reportOverlayContentBounds: (measurement: OverlayContentMeasurement) => void;
  onConfigHotReload: (callback: (payload: ConfigHotReloadPayload) => void) => void;
}

declare global {
  interface Window {
    electronAPI: ElectronAPI;
  }
}

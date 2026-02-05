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
  noun = "noun",
  verb = "verb",
  i_adjective = "i_adjective",
  na_adjective = "na_adjective",
  particle = "particle",
  bound_auxiliary = "bound_auxiliary",
  symbol = "symbol",
  other = "other",
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
  isMerged: boolean;
}

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

export type SecondarySubMode = "hidden" | "visible" | "hover";

export interface SecondarySubConfig {
  secondarySubLanguages?: string[];
  autoLoadSecondarySub?: boolean;
  defaultMode?: SecondarySubMode;
}

export interface WebSocketConfig {
  enabled?: boolean | "auto";
  port?: number;
}

export interface TexthookerConfig {
  openBrowser?: boolean;
}

export interface AnkiConnectConfig {
  enabled?: boolean;
  url?: string;
  pollingRate?: number;
  audioField?: string;
  imageField?: string;
  sentenceField?: string;
  generateAudio?: boolean;
  generateImage?: boolean;
  imageType?: "static" | "avif";
  imageFormat?: "jpg" | "png" | "webp";
  overwriteAudio?: boolean;
  overwriteImage?: boolean;
  mediaInsertMode?: "append" | "prepend";
  audioPadding?: number;
  fallbackDuration?: number;
  deck?: string;
  miscInfoField?: string;
  miscInfoPattern?: string;
  highlightWord?: boolean;
  notificationType?: "osd" | "system" | "both" | "none";
  imageQuality?: number;
  imageMaxWidth?: number;
  imageMaxHeight?: number;
  animatedFps?: number;
  animatedMaxWidth?: number;
  animatedMaxHeight?: number;
  animatedCrf?: number;
  autoUpdateNewCards?: boolean;
  maxMediaDuration?: number;
  sentenceCardModel?: string;
  sentenceCardSentenceField?: string;
  sentenceCardAudioField?: string;
  isLapis?: boolean;
}

export interface SubtitleStyleConfig {
  fontFamily?: string;
  fontSize?: number;
  fontColor?: string;
  fontWeight?: string;
  fontStyle?: string;
  backgroundColor?: string;
  secondary?: {
    fontFamily?: string;
    fontSize?: number;
    fontColor?: string;
    fontWeight?: string;
    fontStyle?: string;
    backgroundColor?: string;
  };
}

export interface ShortcutsConfig {
  copySubtitle?: string | null;
  copySubtitleMultiple?: string | null;
  updateLastCardFromClipboard?: string | null;
  mineSentence?: string | null;
  mineSentenceMultiple?: string | null;
  multiCopyTimeoutMs?: number;
  toggleSecondarySub?: string | null;
}

export type JimakuLanguagePreference = "ja" | "en" | "none";

export interface JimakuConfig {
  apiKey?: string;
  apiKeyCommand?: string;
  apiBaseUrl?: string;
  languagePreference?: JimakuLanguagePreference;
  maxEntryResults?: number;
}

export interface Config {
  subtitlePosition?: SubtitlePosition;
  subtitleFontSize?: number;
  keybindings?: Keybinding[];
  websocket?: WebSocketConfig;
  texthooker?: TexthookerConfig;
  ankiConnect?: AnkiConnectConfig;
  shortcuts?: ShortcutsConfig;
  secondarySub?: SecondarySubConfig;
  subtitleStyle?: SubtitleStyleConfig;
  auto_start_overlay?: boolean;
  jimaku?: JimakuConfig;
}

export interface SubtitleData {
  text: string;
  tokens: MergedToken[] | null;
}

export interface MecabStatus {
  available: boolean;
  enabled: boolean;
  path: string | null;
}

export type JimakuConfidence = "high" | "medium" | "low";

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

export type JimakuApiResponse<T> =
  | { ok: true; data: T }
  | { ok: false; error: JimakuApiError };

export type JimakuDownloadResult =
  | { ok: true; path: string }
  | { ok: false; error: JimakuApiError };

export interface ElectronAPI {
  onSubtitle: (callback: (data: SubtitleData) => void) => void;
  onVisibility: (callback: (visible: boolean) => void) => void;
  onSubtitlePosition: (callback: (position: SubtitlePosition | null) => void) => void;
  getOverlayVisibility: () => Promise<boolean>;
  getCurrentSubtitle: () => Promise<SubtitleData>;
  setIgnoreMouseEvents: (
    ignore: boolean,
    options?: { forward?: boolean },
  ) => void;
  openYomitanSettings: () => void;
  getSubtitlePosition: () => Promise<SubtitlePosition | null>;
  saveSubtitlePosition: (position: SubtitlePosition) => void;
  getMecabStatus: () => Promise<MecabStatus>;
  setMecabEnabled: (enabled: boolean) => void;
  sendMpvCommand: (command: (string | number)[]) => void;
  getKeybindings: () => Promise<Keybinding[]>;
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
  getSubtitleStyle: () => Promise<SubtitleStyleConfig | null>;
}

declare global {
  interface Window {
    electronAPI: ElectronAPI;
  }
}

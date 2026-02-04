/*
 * mpv-yomitan - Yomitan integration for mpv
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

export interface Keybinding {
  key: string;
  command: string[] | null;
}

export interface WebSocketConfig {
  enabled?: boolean | "auto";
  port?: number;
}

export interface TexthookerConfig {
  openBrowser?: boolean;
}

export interface Config {
  subtitlePosition?: SubtitlePosition;
  keybindings?: Keybinding[];
  websocket?: WebSocketConfig;
  texthooker?: TexthookerConfig;
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

export interface ElectronAPI {
  onSubtitle: (callback: (data: SubtitleData) => void) => void;
  onVisibility: (callback: (visible: boolean) => void) => void;
  getOverlayVisibility: () => Promise<boolean>;
  getCurrentSubtitle: () => Promise<SubtitleData>;
  setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => void;
  openYomitanSettings: () => void;
  getSubtitlePosition: () => Promise<SubtitlePosition | null>;
  saveSubtitlePosition: (position: SubtitlePosition) => void;
  getMecabStatus: () => Promise<MecabStatus>;
  setMecabEnabled: (enabled: boolean) => void;
  sendMpvCommand: (command: string[]) => void;
  getKeybindings: () => Promise<Keybinding[]>;
  quitApp: () => void;
  toggleDevTools: () => void;
  toggleOverlay: () => void;
}

declare global {
  interface Window {
    electronAPI: ElectronAPI;
  }
}

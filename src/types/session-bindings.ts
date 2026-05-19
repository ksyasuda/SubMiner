export type SessionKeyModifier = 'ctrl' | 'alt' | 'shift' | 'meta';

export type SessionActionId =
  | 'toggleStatsOverlay'
  | 'markWatched'
  | 'toggleVisibleOverlay'
  | 'copySubtitle'
  | 'copySubtitleMultiple'
  | 'updateLastCardFromClipboard'
  | 'triggerFieldGrouping'
  | 'triggerSubsync'
  | 'mineSentence'
  | 'mineSentenceMultiple'
  | 'toggleSecondarySub'
  | 'toggleSubtitleSidebar'
  | 'markAudioCard'
  | 'openRuntimeOptions'
  | 'openSessionHelp'
  | 'openCharacterDictionary'
  | 'openControllerSelect'
  | 'openControllerDebug'
  | 'openJimaku'
  | 'openYoutubePicker'
  | 'openPlaylistBrowser'
  | 'replayCurrentSubtitle'
  | 'playNextSubtitle'
  | 'shiftSubDelayPrevLine'
  | 'shiftSubDelayNextLine'
  | 'cycleRuntimeOption';

export interface SessionKeySpec {
  code: string;
  modifiers: SessionKeyModifier[];
}

export interface SessionBindingWarning {
  kind: 'unsupported' | 'conflict' | 'deprecated-config';
  path: string;
  message: string;
  value: unknown;
  conflictingPaths?: string[];
}

export interface SessionActionPayload {
  count?: number;
  runtimeOptionId?: string;
  direction?: 1 | -1;
}

type CompiledSessionBindingBase = {
  sourcePath: string;
  originalKey: string;
  key: SessionKeySpec;
};

export interface CompiledMpvCommandBinding extends CompiledSessionBindingBase {
  actionType: 'mpv-command';
  command: (string | number)[];
}

export interface CompiledSessionActionBinding extends CompiledSessionBindingBase {
  actionType: 'session-action';
  actionId: SessionActionId;
  payload?: SessionActionPayload;
}

export type CompiledSessionBinding = CompiledMpvCommandBinding | CompiledSessionActionBinding;

export interface PluginSessionBindingsArtifact {
  version: 1;
  generatedAt: string;
  numericSelectionTimeoutMs: number;
  bindings: CompiledSessionBinding[];
  warnings: SessionBindingWarning[];
}
